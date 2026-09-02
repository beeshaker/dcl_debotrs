# -*- coding: utf-8 -*-
import base64
import io
import math
import os
import tempfile
from collections import defaultdict
from datetime import date, datetime

from odoo import api, fields, models, _
from odoo.exceptions import UserError, ValidationError

try:
    import openpyxl
except ImportError:
    openpyxl = None

try:
    import xlsxwriter
except ImportError:
    xlsxwriter = None

from ..models.aging_engine import (
    calculate_debtor_ageing,
    normalize_header,
    parse_month,
)


class DclDebtorsAgeWizard(models.TransientModel):
    _name = "dcl.debtors.age.wizard"
    _description = "DCL Debtors Age Analysis"

    source_file = fields.Binary(string="Source Excel", required=True)
    source_filename = fields.Char(string="Source Filename")
    reporting_date = fields.Date(
        string="Reporting Date",
        required=True,
        default=fields.Date.context_today,
        help="The 'as at' date for the age analysis. The month/year is used for the current ageing period.",
    )
    source_sheet = fields.Char(
        string="Source Sheet",
        help="Leave blank to auto-detect the worksheet containing the monthly debtors ledger.",
    )

    result_file = fields.Binary(string="Generated Excel", readonly=True)
    result_filename = fields.Char(string="Generated Filename", readonly=True)

    property_count = fields.Integer(string="Properties", readonly=True)
    debtor_count = fields.Integer(string="Debtors", readonly=True)
    exception_count = fields.Integer(string="Exceptions", readonly=True)
    total_receivables = fields.Monetary(string="Total Receivables", readonly=True)
    status_message = fields.Text(string="Status", readonly=True)
    currency_id = fields.Many2one(
        "res.currency",
        string="Currency",
        default=lambda self: self.env.company.currency_id,
        readonly=True,
    )

    # -------------------------------------------------------------------------
    # Public actions
    # -------------------------------------------------------------------------

    def action_generate(self):
        self.ensure_one()

        if not self.source_file:
            raise ValidationError(_("Please upload the source Debtors List Excel file."))

        if openpyxl is None:
            raise UserError(_("Python dependency 'openpyxl' is not installed."))
        if xlsxwriter is None:
            raise UserError(_("Python dependency 'xlsxwriter' is not installed."))

        try:
            workbook = self._load_workbook()
            worksheet = self._select_sheet(workbook)

            parsed = self._read_source(worksheet)

            if not parsed["rows"]:
                raise UserError(
                    _("No debtor rows were found in worksheet '%s'.") % worksheet.title
                )

            result_bytes = self._build_xlsx(parsed)

            total_receivables = sum(
                float(row.get("current_outstanding") or 0.0)
                for row in parsed["rows"]
            )

            exception_count = sum(
                1 for row in parsed["rows"] if row.get("exception")
            )

            property_names = {
                row.get("property") or "Unassigned"
                for row in parsed["rows"]
            }

            filename = "Debtors_Age_Analysis_%s.xlsx" % (
                self.reporting_date.strftime("%Y-%m-%d")
                if self.reporting_date
                else fields.Date.today().strftime("%Y-%m-%d")
            )

            self.write({
                "result_file": base64.b64encode(result_bytes),
                "result_filename": filename,
                "property_count": len(property_names),
                "debtor_count": len(parsed["rows"]),
                "exception_count": exception_count,
                "total_receivables": total_receivables,
                "status_message": _(
                    "Age analysis generated successfully from worksheet '%s'. "
                    "%s debtors across %s properties were processed."
                ) % (
                    worksheet.title,
                    len(parsed["rows"]),
                    len(property_names),
                ),
            })

            return {
                "type": "ir.actions.act_window",
                "name": _("Debtors Age Analysis"),
                "res_model": self._name,
                "res_id": self.id,
                "view_mode": "form",
                "target": "new",
            }

        except UserError:
            raise
        except ValidationError:
            raise
        except Exception as exc:
            raise UserError(
                _("Unable to generate the age analysis.\n\n%s") % str(exc)
            ) from exc

    def action_download(self):
        self.ensure_one()

        if not self.result_file:
            raise UserError(_("Generate the age analysis before downloading."))

        return {
            "type": "ir.actions.act_url",
            "url": (
                "/web/content?"
                "model=%s&id=%s&field=result_file&filename_field=result_filename"
                "&download=true"
            ) % (self._name, self.id),
            "target": "self",
        }

    # -------------------------------------------------------------------------
    # Workbook loading / source discovery
    # -------------------------------------------------------------------------

    def _load_workbook(self):
        payload = base64.b64decode(self.source_file)
        return openpyxl.load_workbook(
            io.BytesIO(payload),
            read_only=True,
            data_only=True,
        )

    @staticmethod
    def _row_values(ws, row_number):
        """
        Read one row sequentially.

        With openpyxl read_only=True we intentionally avoid thousands of
        ws.cell(row, col) calls. Sequential iteration is dramatically faster.
        """
        for row in ws.iter_rows(
            min_row=row_number,
            max_row=row_number,
            values_only=True,
        ):
            return tuple(row)
        return tuple()

    def _score_sheet(self, ws):
        """
        Score a sheet based on month headers / ledger labels in its first rows.
        This only touches the first few rows, so it remains cheap.
        """
        first = self._row_values(ws, 1)
        second = self._row_values(ws, 2)

        score = 0

        for value in first:
            if parse_month(value):
                score += 4

        expected = {
            "arrears b f",
            "rent levy",
            "recoveries",
            "adjustments",
            "receipts",
            "current bal",
            "opening balance",
            "total billings",
            "total receipts",
            "closing balance",
        }

        for value in list(first) + list(second):
            header = normalize_header(value)
            if header in expected:
                score += 2

        score += min(ws.max_column or 0, 100) / 100.0
        score += min(ws.max_row or 0, 10000) / 10000.0
        return score

    def _select_sheet(self, workbook):
        requested = (self.source_sheet or "").strip()

        if requested:
            if requested not in workbook.sheetnames:
                raise UserError(
                    _("Worksheet '%s' does not exist. Available sheets: %s")
                    % (requested, ", ".join(workbook.sheetnames))
                )
            return workbook[requested]

        candidates = [(self._score_sheet(ws), ws) for ws in workbook.worksheets]
        candidates.sort(key=lambda item: item[0], reverse=True)

        if not candidates or candidates[0][0] <= 0:
            raise UserError(_("Could not identify the debtors ledger worksheet."))

        return candidates[0][1]

    # -------------------------------------------------------------------------
    # Source parser
    # -------------------------------------------------------------------------

    @staticmethod
    def _safe_number(value):
        if value in (None, ""):
            return 0.0

        if isinstance(value, bool):
            return float(value)

        if isinstance(value, (int, float)):
            if isinstance(value, float) and math.isnan(value):
                return 0.0
            return float(value)

        text = str(value).strip()
        if not text:
            return 0.0

        text = text.replace(",", "")
        text = text.replace("(", "-").replace(")", "")
        text = text.replace("KES", "").strip()

        try:
            return float(text)
        except (TypeError, ValueError):
            return 0.0

    @staticmethod
    def _text(value):
        if value is None:
            return ""
        return str(value).strip()

    def _build_column_map(self, header1, header2):
        """
        Build all column indexes once from the two header rows.

        Returns:
            {
                "months": {
                    date(2026, 8, 1): {
                        "arrears": 3,
                        "rent": 4,
                        ...
                    }
                },
                "summary": {
                    "opening_balance": 52,
                    ...
                }
            }

        Indexes are zero-based because rows are tuples.
        """
        months = {}
        summary = {}

        current_month = None

        field_aliases = {
            "arrears b f": "arrears",
            "arrears bf": "arrears",
            "arrears b/f": "arrears",
            "rent levy": "rent_levy",
            "rent / levy": "rent_levy",
            "recoveries": "recoveries",
            "adjustments": "adjustments",
            "receipts": "receipts",
            "current bal": "current_balance",
            "current balance": "current_balance",
        }

        summary_aliases = {
            "opening balance": "opening_balance",
            "total billings": "total_billings",
            "total receipts": "total_receipts",
            "closing balance": "closing_balance",
        }

        width = max(len(header1), len(header2))

        for idx in range(width):
            top = header1[idx] if idx < len(header1) else None
            sub = header2[idx] if idx < len(header2) else None

            parsed_month = parse_month(top)
            if parsed_month:
                current_month = parsed_month.replace(day=1)

            top_norm = normalize_header(top)
            sub_norm = normalize_header(sub)

            if top_norm in summary_aliases:
                summary[summary_aliases[top_norm]] = idx
            if sub_norm in summary_aliases:
                summary[summary_aliases[sub_norm]] = idx

            field_name = field_aliases.get(sub_norm) or field_aliases.get(top_norm)

            if current_month and field_name:
                months.setdefault(current_month, {})[field_name] = idx

        return {
            "months": dict(sorted(months.items())),
            "summary": summary,
        }

    def _detect_identity_columns(self, header1, header2):
        """
        Detect the important non-month columns conservatively.
        Falls back to the known source layout if headings are not explicit.
        """
        combined = []
        width = max(len(header1), len(header2))

        for idx in range(width):
            text = " ".join(
                part
                for part in [
                    normalize_header(header1[idx] if idx < len(header1) else None),
                    normalize_header(header2[idx] if idx < len(header2) else None),
                ]
                if part
            )
            combined.append(text)

        def find_index(words, default):
            for idx, text in enumerate(combined):
                if any(word in text for word in words):
                    return idx
            return default

        # Source workbook is known to use account/debtor ID in column A.
        account_idx = find_index(
            ["debtor id", "account id", "customer id", "tenant id", "code"],
            0,
        )
        name_idx = find_index(
            ["debtor name", "customer name", "tenant name", "client name", "name"],
            1,
        )
        property_idx = find_index(
            ["property", "site", "building", "estate", "project"],
            2,
        )

        return {
            "account_id": account_idx,
            "name": name_idx,
            "property": property_idx,
        }

    def _read_source(self, ws):
        """
        Fast source reader.

        The key performance change is that the worksheet is consumed once with
        iter_rows(values_only=True). We do not use ws.cell() inside the debtor
        loop.
        """
        row_iterator = ws.iter_rows(values_only=True)

        try:
            header1 = tuple(next(row_iterator))
        except StopIteration:
            raise UserError(_("The selected worksheet is empty."))

        try:
            header2 = tuple(next(row_iterator))
        except StopIteration:
            header2 = tuple()

        column_map = self._build_column_map(header1, header2)
        month_blocks = column_map["months"]
        summary_columns = column_map["summary"]

        if not month_blocks:
            raise UserError(
                _("No monthly ledger blocks could be detected in worksheet '%s'.")
                % ws.title
            )

        identity = self._detect_identity_columns(header1, header2)

        rows = []

        def cell(row_tuple, idx):
            if idx is None or idx < 0 or idx >= len(row_tuple):
                return None
            return row_tuple[idx]

        for row_tuple in row_iterator:
            account_id = self._text(cell(row_tuple, identity["account_id"]))
            debtor_name = self._text(cell(row_tuple, identity["name"]))
            property_name = self._text(cell(row_tuple, identity["property"]))

            # Skip headings/totals/blank lines.
            if not account_id and not debtor_name:
                continue

            joined_identity = " ".join(
                [account_id.lower(), debtor_name.lower(), property_name.lower()]
            )
            if any(
                marker in joined_identity
                for marker in (
                    "grand total",
                    "portfolio total",
                    "property total",
                    "subtotal",
                )
            ):
                continue

            monthly_rows = []

            for month, cols in month_blocks.items():
                monthly_rows.append({
                    "month": month,
                    "arrears_bf": self._safe_number(
                        cell(row_tuple, cols.get("arrears"))
                    ),
                    "rent_levy": self._safe_number(
                        cell(row_tuple, cols.get("rent_levy"))
                    ),
                    "recoveries": self._safe_number(
                        cell(row_tuple, cols.get("recoveries"))
                    ),
                    "adjustments": self._safe_number(
                        cell(row_tuple, cols.get("adjustments"))
                    ),
                    "receipts": self._safe_number(
                        cell(row_tuple, cols.get("receipts"))
                    ),
                    "current_balance": self._safe_number(
                        cell(row_tuple, cols.get("current_balance"))
                    ),
                })

            source_opening = self._safe_number(
                cell(row_tuple, summary_columns.get("opening_balance"))
            )
            source_billings = self._safe_number(
                cell(row_tuple, summary_columns.get("total_billings"))
            )
            source_receipts = self._safe_number(
                cell(row_tuple, summary_columns.get("total_receipts"))
            )
            source_closing = self._safe_number(
                cell(row_tuple, summary_columns.get("closing_balance"))
            )

            # If summary columns were unavailable, use the latest period balance.
            if "closing_balance" not in summary_columns and monthly_rows:
                source_closing = monthly_rows[-1]["current_balance"]

            calculation = calculate_debtor_ageing(
                opening_balance=source_opening,
                monthly_rows=monthly_rows,
                source_closing_balance=source_closing,
            )

            rows.append({
                "account_id": account_id,
                "debtor_name": debtor_name,
                "property": property_name or "Unassigned",
                "billing_frequency": calculation.get("billing_frequency", ""),
                "opening_balance": source_opening,
                "total_billings": source_billings,
                "total_receipts": source_receipts,
                "current_outstanding": source_closing,
                "ageing": calculation.get("ageing", {}),
                "reconciliation_difference": calculation.get(
                    "reconciliation_difference", 0.0
                ),
                "exception": bool(calculation.get("exception")),
            })

        return {
            "rows": rows,
            "months": list(month_blocks.keys()),
            "sheet_name": ws.title,
        }

    # -------------------------------------------------------------------------
    # XLSX generation
    # -------------------------------------------------------------------------

    @staticmethod
    def _safe_sheet_name(name, used):
        invalid = '[]:*?/\\'
        cleaned = "".join("_" if char in invalid else char for char in (name or "Property"))
        cleaned = cleaned.strip() or "Property"
        cleaned = cleaned[:31]

        base = cleaned
        suffix = 1

        while cleaned.lower() in used:
            suffix += 1
            tag = "_%s" % suffix
            cleaned = (base[: 31 - len(tag)] + tag)[:31]

        used.add(cleaned.lower())
        return cleaned

    def _build_xlsx(self, parsed):
        """
        Build the XLSX on disk rather than entirely in RAM.

        This is intentionally file-backed for a small VPS. The generated bytes
        are only loaded after xlsxwriter has closed the workbook.
        """
        rows = parsed["rows"]
        source_months = sorted(parsed["months"], reverse=True)

        grouped = defaultdict(list)
        for row in rows:
            grouped[row.get("property") or "Unassigned"].append(row)

        fd, path = tempfile.mkstemp(prefix="dcl_debtors_age_", suffix=".xlsx")
        os.close(fd)

        try:
            workbook = xlsxwriter.Workbook(path)

            fmt_title = workbook.add_format({
                "bold": True,
                "font_size": 16,
            })
            fmt_header = workbook.add_format({
                "bold": True,
                "border": 1,
                "text_wrap": True,
                "valign": "vcenter",
            })
            fmt_money = workbook.add_format({
                "num_format": '#,##0.00;[Red]-#,##0.00',
            })
            fmt_money_exception = workbook.add_format({
                "num_format": '#,##0.00;[Red]-#,##0.00',
                "bg_color": "#C6EFCE",
            })
            fmt_text_exception = workbook.add_format({
                "bg_color": "#C6EFCE",
            })
            fmt_date = workbook.add_format({
                "num_format": "dd-mmm-yyyy",
            })

            # ---------------- Portfolio Summary ----------------
            summary = workbook.add_worksheet("Portfolio Summary")
            summary.freeze_panes(4, 0)

            summary.write(0, 0, "DCL Debtors Age Analysis", fmt_title)
            summary.write(1, 0, "Reporting Date")
            if self.reporting_date:
                summary.write_datetime(
                    1,
                    1,
                    datetime.combine(self.reporting_date, datetime.min.time()),
                    fmt_date,
                )

            summary_headers = [
                "Property",
                "Debtors",
                "Current Outstanding",
                "Exceptions",
            ]
            for col, heading in enumerate(summary_headers):
                summary.write(3, col, heading, fmt_header)

            summary_row = 4
            for property_name in sorted(grouped, key=lambda x: x.lower()):
                property_rows = grouped[property_name]
                total = sum(
                    float(item.get("current_outstanding") or 0.0)
                    for item in property_rows
                )
                exceptions = sum(
                    1 for item in property_rows if item.get("exception")
                )

                summary.write(summary_row, 0, property_name)
                summary.write_number(summary_row, 1, len(property_rows))
                summary.write_number(summary_row, 2, total, fmt_money)
                summary.write_number(summary_row, 3, exceptions)
                summary_row += 1

            summary.write(summary_row, 0, "TOTAL", fmt_header)
            summary.write_number(summary_row, 1, len(rows), fmt_header)
            summary.write_number(
                summary_row,
                2,
                sum(float(item.get("current_outstanding") or 0.0) for item in rows),
                fmt_money,
            )
            summary.write_number(
                summary_row,
                3,
                sum(1 for item in rows if item.get("exception")),
                fmt_header,
            )

            summary.set_column(0, 0, 32)
            summary.set_column(1, 1, 12)
            summary.set_column(2, 2, 20)
            summary.set_column(3, 3, 12)

            # ---------------- Property sheets ----------------
            used_sheet_names = {"portfolio summary"}

            for property_name in sorted(grouped, key=lambda x: x.lower()):
                property_rows = grouped[property_name]

                sheet_name = self._safe_sheet_name(property_name, used_sheet_names)
                ws = workbook.add_worksheet(sheet_name)
                ws.freeze_panes(3, 3)
                ws.autofilter(
                    2,
                    0,
                    2 + max(len(property_rows), 1),
                    8 + len(source_months),
                )

                ws.write(0, 0, property_name, fmt_title)

                headers = [
                    "Account ID",
                    "Debtor Name",
                    "Billing Frequency",
                    "Opening Balance",
                    "Total Billings",
                    "Total Receipts",
                    "Current Outstanding",
                ]

                month_headers = [
                    month.strftime("%b %Y") for month in source_months
                ]

                headers += month_headers
                headers += [
                    "Opening / Older",
                    "Reconciliation Difference",
                    "Exception",
                ]

                for col, heading in enumerate(headers):
                    ws.write(2, col, heading, fmt_header)

                out_row = 3

                for debtor in property_rows:
                    exception = debtor.get("exception")
                    text_fmt = fmt_text_exception if exception else None
                    money_fmt = fmt_money_exception if exception else fmt_money

                    ws.write(
                        out_row,
                        0,
                        debtor.get("account_id", ""),
                        text_fmt,
                    )
                    ws.write(
                        out_row,
                        1,
                        debtor.get("debtor_name", ""),
                        text_fmt,
                    )
                    ws.write(
                        out_row,
                        2,
                        debtor.get("billing_frequency", ""),
                        text_fmt,
                    )

                    ws.write_number(
                        out_row,
                        3,
                        float(debtor.get("opening_balance") or 0.0),
                        money_fmt,
                    )
                    ws.write_number(
                        out_row,
                        4,
                        float(debtor.get("total_billings") or 0.0),
                        money_fmt,
                    )
                    ws.write_number(
                        out_row,
                        5,
                        float(debtor.get("total_receipts") or 0.0),
                        money_fmt,
                    )
                    ws.write_number(
                        out_row,
                        6,
                        float(debtor.get("current_outstanding") or 0.0),
                        money_fmt,
                    )

                    ageing = debtor.get("ageing") or {}

                    col = 7
                    for month in source_months:
                        value = float(ageing.get(month, 0.0) or 0.0)
                        ws.write_number(out_row, col, value, money_fmt)
                        col += 1

                    older_value = float(
                        ageing.get("opening", ageing.get("older", 0.0)) or 0.0
                    )
                    ws.write_number(out_row, col, older_value, money_fmt)
                    col += 1

                    ws.write_number(
                        out_row,
                        col,
                        float(debtor.get("reconciliation_difference") or 0.0),
                        money_fmt,
                    )
                    col += 1

                    ws.write(
                        out_row,
                        col,
                        "Yes" if exception else "",
                        text_fmt,
                    )

                    out_row += 1

                # Total row
                ws.write(out_row, 0, "TOTAL", fmt_header)

                for col in range(3, 7 + len(source_months) + 2):
                    if out_row > 3:
                        first = xlsxwriter.utility.xl_rowcol_to_cell(3, col)
                        last = xlsxwriter.utility.xl_rowcol_to_cell(out_row - 1, col)
                        ws.write_formula(
                            out_row,
                            col,
                            "=SUM(%s:%s)" % (first, last),
                            fmt_money,
                        )

                ws.set_column(0, 0, 18)
                ws.set_column(1, 1, 34)
                ws.set_column(2, 2, 18)
                ws.set_column(3, len(headers) - 1, 16)

            workbook.close()

            with open(path, "rb") as handle:
                return handle.read()

        finally:
            try:
                if os.path.exists(path):
                    os.remove(path)
            except OSError:
                pass
