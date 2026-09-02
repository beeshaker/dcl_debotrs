import base64
from collections import defaultdict
from datetime import date
from io import BytesIO
import logging

from odoo import _, api, fields, models
from odoo.exceptions import UserError, ValidationError

from ..models.aging_engine import (
    as_number,
    calculate_debtor_ageing,
    normalize_header,
    parse_month,
)

_logger = logging.getLogger(__name__)


class DebtorsAgeWizard(models.TransientModel):
    _name = "dcl.debtors.age.wizard"
    _description = "Debtors Age Analysis Generator"

    source_file = fields.Binary(string="Debtors List (.xlsx)", required=True, attachment=True)
    source_filename = fields.Char(string="Source Filename")
    reporting_date = fields.Date(string="Reporting Date", required=True, default=fields.Date.context_today)
    source_sheet = fields.Char(
        string="Source Sheet",
        help="Optional. Leave blank to automatically select the sheet with the richest debtor data.",
    )

    result_file = fields.Binary(string="Generated Age Analysis", readonly=True, attachment=True)
    result_filename = fields.Char(readonly=True)
    status_message = fields.Text(readonly=True)
    property_count = fields.Integer(readonly=True)
    debtor_count = fields.Integer(readonly=True)
    exception_count = fields.Integer(readonly=True)
    total_receivables = fields.Monetary(readonly=True, currency_field="currency_id")
    currency_id = fields.Many2one(
        "res.currency",
        default=lambda self: self.env.company.currency_id,
        readonly=True,
    )

    def _load_workbook(self):
        self.ensure_one()
        if not self.source_file:
            raise ValidationError(_("Please upload the Debtors List workbook."))
        try:
            from openpyxl import load_workbook
        except ImportError as exc:
            raise UserError(_("Python package 'openpyxl' is required on the Odoo server.")) from exc
        try:
            payload = base64.b64decode(self.source_file)
            return load_workbook(BytesIO(payload), read_only=True, data_only=True)
        except Exception as exc:
            _logger.exception("Unable to read debtors workbook")
            raise UserError(_("The uploaded file could not be read as an .xlsx workbook: %s") % exc) from exc

    def _score_sheet(self, ws):
        month_headers = 0
        for col in range(1, min(ws.max_column, 150) + 1):
            if normalize_header(ws.cell(1, col).value) == "arrears_bf" and parse_month(
                ws.cell(2, col).value, self.reporting_date.year
            ):
                month_headers += 1
        return (month_headers * 100000) + ws.max_row

    def _select_sheet(self, workbook):
        if self.source_sheet:
            if self.source_sheet not in workbook.sheetnames:
                raise UserError(_("Sheet '%s' was not found in the workbook.") % self.source_sheet)
            return workbook[self.source_sheet]
        candidates = sorted(workbook.worksheets, key=self._score_sheet, reverse=True)
        if not candidates or self._score_sheet(candidates[0]) < 100000:
            raise UserError(_("No valid monthly debtors ledger sheet was detected."))
        return candidates[0]

    def _detect_month_blocks(self, ws):
        blocks = {}
        current_month = None
        for col in range(1, ws.max_column + 1):
            header = normalize_header(ws.cell(1, col).value)
            month_value = parse_month(ws.cell(2, col).value, self.reporting_date.year)
            if month_value:
                current_month = month_value
            if current_month and header in {
                "arrears_bf", "rent_levy", "recoveries", "adjustments", "receipts", "current_bal"
            }:
                blocks.setdefault(current_month, {})[header] = col
        valid = {
            month: mapping
            for month, mapping in blocks.items()
            if "arrears_bf" in mapping and "current_bal" in mapping
        }
        if not valid:
            raise UserError(_("Could not detect any complete monthly blocks in rows 1-2."))
        return dict(sorted(valid.items()))

    def _find_summary_columns(self, ws):
        cols = {}
        for col in range(1, ws.max_column + 1):
            h = str(ws.cell(1, col).value or "").strip().lower()
            if h == "opening balance":
                cols["opening_balance"] = col
            elif h == "total billings":
                cols["total_billings"] = col
            elif h == "total receipts":
                cols["total_receipts"] = col
            elif h == "closing balance":
                cols["closing_balance"] = col
        return cols

    def _read_source(self, ws):
        month_blocks = self._detect_month_blocks(ws)
        summary_cols = self._find_summary_columns(ws)
        results = []
        skipped = 0

        for row_idx in range(3, ws.max_row + 1):
            debtor_id = ws.cell(row_idx, 1).value
            debtor_name = ws.cell(row_idx, 2).value
            property_name = ws.cell(row_idx, 3).value
            if not debtor_id and not debtor_name and not property_name:
                continue
            if not debtor_name or not property_name:
                skipped += 1
                continue

            monthly_rows = []
            for month, mapping in month_blocks.items():
                monthly_rows.append({
                    "month": month,
                    "arrears_bf": as_number(ws.cell(row_idx, mapping.get("arrears_bf")).value),
                    "rent_levy": as_number(ws.cell(row_idx, mapping.get("rent_levy")).value) if mapping.get("rent_levy") else 0.0,
                    "recoveries": as_number(ws.cell(row_idx, mapping.get("recoveries")).value) if mapping.get("recoveries") else 0.0,
                    "adjustments": as_number(ws.cell(row_idx, mapping.get("adjustments")).value) if mapping.get("adjustments") else 0.0,
                    "receipts": as_number(ws.cell(row_idx, mapping.get("receipts")).value) if mapping.get("receipts") else 0.0,
                    "current_bal": as_number(ws.cell(row_idx, mapping.get("current_bal")).value),
                })

            latest = max(month_blocks)
            source_closing = monthly_rows[-1]["current_bal"]
            if summary_cols.get("closing_balance"):
                source_closing = as_number(ws.cell(row_idx, summary_cols["closing_balance"]).value)

            results.append(calculate_debtor_ageing(
                debtor_id=debtor_id,
                debtor_name=debtor_name,
                property_name=property_name,
                monthly_rows=monthly_rows,
                source_closing=source_closing,
            ))

        if not results:
            raise UserError(_("No debtor rows were found in the selected sheet."))
        return results, month_blocks, skipped

    @staticmethod
    def _safe_sheet_name(name, used):
        invalid = set('[]:*?/\\')
        cleaned = ''.join('_' if c in invalid else c for c in (name or "Property")).strip()[:31] or "Property"
        base = cleaned
        i = 2
        while cleaned.lower() in used:
            suffix = " %s" % i
            cleaned = (base[:31-len(suffix)] + suffix)
            i += 1
        used.add(cleaned.lower())
        return cleaned

    def _build_xlsx(self, results, month_blocks):
        try:
            import xlsxwriter
        except ImportError as exc:
            raise UserError(_("Python package 'xlsxwriter' is required on the Odoo server.")) from exc

        output = BytesIO()
        wb = xlsxwriter.Workbook(output, {"in_memory": True})

        title_fmt = wb.add_format({"bold": True, "font_size": 16, "align": "left"})
        subtitle_fmt = wb.add_format({"font_size": 10, "font_color": "#555555"})
        header_fmt = wb.add_format({"bold": True, "bg_color": "#D9EAF7", "border": 1, "text_wrap": True, "valign": "vcenter"})
        money_fmt = wb.add_format({"num_format": '#,##0.00;[Red]-#,##0.00', "border": 1})
        text_fmt = wb.add_format({"border": 1})
        int_fmt = wb.add_format({"num_format": '0', "border": 1})
        total_fmt = wb.add_format({"bold": True, "num_format": '#,##0.00;[Red]-#,##0.00', "top": 1})
        total_text_fmt = wb.add_format({"bold": True, "top": 1})
        exception_fmt = wb.add_format({"bg_color": "#C6EFCE", "font_color": "#006100", "border": 1, "text_wrap": True})
        kpi_fmt = wb.add_format({"bold": True, "font_size": 12, "bg_color": "#EAF2F8", "border": 1, "align": "center", "valign": "vcenter"})

        months_desc = sorted(month_blocks.keys(), reverse=True)
        month_labels = [m.strftime("%b %Y") for m in months_desc]
        properties = defaultdict(list)
        for result in results:
            properties[result.property_name].append(result)

        summary = wb.add_worksheet("Portfolio Summary")
        summary.hide_gridlines(2)
        summary.merge_range(0, 0, 0, 9, "PORTFOLIO DEBTORS AGING & RECONCILIATION MASTER SUMMARY", title_fmt)
        summary.merge_range(
            1, 0, 1, 9,
            "Reporting Date: %s  |  Currency: %s  |  Properties: %s" % (
                self.reporting_date.strftime("%d %B %Y"), self.currency_id.name, len(properties)
            ), subtitle_fmt
        )
        total_receivables = sum(r.current_outstanding for r in results)
        total_charges = sum(r.period_charges for r in results)
        total_receipts = -sum(r.period_receipts for r in results)
        exceptions = sum(bool(r.exception) for r in results)
        summary.merge_range(3, 0, 4, 1, "Total Receivables\n%s %s" % (self.currency_id.name, f"{total_receivables:,.2f}"), kpi_fmt)
        summary.merge_range(3, 3, 4, 4, "Monitored Sites\n%s Properties" % len(properties), kpi_fmt)
        summary.merge_range(3, 6, 4, 7, "Total Debtors\n%s" % len(results), kpi_fmt)
        summary.merge_range(3, 9, 4, 10, "Exceptions\n%s" % exceptions, kpi_fmt)

        headers = ["#", "Property / Site Name", "Debtors", "Opening Arrears", "Period Charges", "Period Receipts", "Current Outstanding"] + month_labels + ["Opening", "Exceptions"]
        header_row = 6
        for col, label in enumerate(headers):
            summary.write(header_row, col, label, header_fmt)

        used_names = {"portfolio summary"}
        for idx, (property_name, debtors) in enumerate(sorted(properties.items()), start=1):
            row = header_row + idx
            summary.write_number(row, 0, idx, int_fmt)
            summary.write(row, 1, property_name, text_fmt)
            summary.write_number(row, 2, len(debtors), int_fmt)
            summary.write_number(row, 3, sum(r.opening_balance for r in debtors), money_fmt)
            summary.write_number(row, 4, sum(r.period_charges for r in debtors), money_fmt)
            summary.write_number(row, 5, -sum(r.period_receipts for r in debtors), money_fmt)
            summary.write_number(row, 6, sum(r.current_outstanding for r in debtors), money_fmt)
            c = 7
            for label in month_labels:
                summary.write_number(row, c, sum(max(r.aged.get(label, 0.0), 0.0) for r in debtors), money_fmt)
                c += 1
            summary.write_number(row, c, sum(max(r.aged.get("Opening", 0.0), 0.0) for r in debtors), money_fmt)
            summary.write_number(row, c + 1, sum(bool(r.exception) for r in debtors), exception_fmt if any(r.exception for r in debtors) else int_fmt)

            sheet_name = self._safe_sheet_name(property_name, used_names)
            ws = wb.add_worksheet(sheet_name)
            ws.hide_gridlines(2)
            ws.freeze_panes(6, 0)
            ws.merge_range(0, 0, 0, min(len(headers), 14), "%s - DEBTORS AGE ANALYSIS" % property_name.upper(), title_fmt)
            ws.write(1, 0, "Reporting Date: %s" % self.reporting_date.strftime("%d %B %Y"), subtitle_fmt)
            detail_headers = ["Debtor ID", "Debtor / Tenant", "Billing Frequency", "Opening Balance", "Period Charges", "Period Receipts", "Current Outstanding"] + month_labels + ["Opening", "Exception"]
            for col, label in enumerate(detail_headers):
                ws.write(5, col, label, header_fmt)

            for ridx, debtor in enumerate(sorted(debtors, key=lambda d: (d.debtor_name.lower(), d.debtor_id)), start=6):
                ws.write(ridx, 0, debtor.debtor_id, text_fmt)
                ws.write(ridx, 1, debtor.debtor_name, text_fmt)
                ws.write(ridx, 2, debtor.billing_frequency, text_fmt)
                ws.write_number(ridx, 3, debtor.opening_balance, money_fmt)
                ws.write_number(ridx, 4, debtor.period_charges, money_fmt)
                ws.write_number(ridx, 5, -debtor.period_receipts, money_fmt)
                ws.write_number(ridx, 6, debtor.current_outstanding, money_fmt)
                c2 = 7
                for label in month_labels:
                    ws.write_number(ridx, c2, max(debtor.aged.get(label, 0.0), 0.0), money_fmt)
                    c2 += 1
                ws.write_number(ridx, c2, max(debtor.aged.get("Opening", 0.0), 0.0), money_fmt)
                ws.write(ridx, c2 + 1, debtor.exception, exception_fmt if debtor.exception else text_fmt)

            total_row = 6 + len(debtors)
            ws.write(total_row, 1, "TOTAL", total_text_fmt)
            for col in range(3, len(detail_headers) - 1):
                start = 7
                end = total_row
                col_letter = xlsxwriter.utility.xl_col_to_name(col)
                ws.write_formula(total_row, col, f"=SUM({col_letter}{start}:{col_letter}{end})", total_fmt)
            ws.set_column(0, 0, 14)
            ws.set_column(1, 1, 38)
            ws.set_column(2, 2, 18)
            ws.set_column(3, len(detail_headers)-2, 15)
            ws.set_column(len(detail_headers)-1, len(detail_headers)-1, 46)
            ws.autofilter(5, 0, total_row - 1, len(detail_headers) - 1)

        summary.set_column(0, 0, 5)
        summary.set_column(1, 1, 30)
        summary.set_column(2, 2, 10)
        summary.set_column(3, len(headers)-2, 16)
        summary.set_column(len(headers)-1, len(headers)-1, 12)
        summary.freeze_panes(header_row + 1, 0)
        summary.autofilter(header_row, 0, header_row + len(properties), len(headers)-1)

        wb.close()
        return output.getvalue()

    def action_generate(self):
        self.ensure_one()
        workbook = self._load_workbook()
        ws = self._select_sheet(workbook)
        results, month_blocks, skipped = self._read_source(ws)
        payload = self._build_xlsx(results, month_blocks)

        properties = {r.property_name for r in results}
        exceptions = sum(bool(r.exception) for r in results)
        self.write({
            "result_file": base64.b64encode(payload),
            "result_filename": "Debtors_Age_Analysis_%s.xlsx" % self.reporting_date.strftime("%Y-%m-%d"),
            "property_count": len(properties),
            "debtor_count": len(results),
            "exception_count": exceptions,
            "total_receivables": sum(r.current_outstanding for r in results),
            "status_message": _(
                "Generated from sheet '%s'. %s properties, %s debtors, %s reconciliation exceptions. %s incomplete rows skipped."
            ) % (ws.title, len(properties), len(results), exceptions, skipped),
        })
        return {
            "type": "ir.actions.act_window",
            "res_model": self._name,
            "res_id": self.id,
            "view_mode": "form",
            "target": "current",
        }

    def action_download(self):
        self.ensure_one()
        if not self.result_file:
            raise UserError(_("Generate the Age Analysis first."))
        return {
            "type": "ir.actions.act_url",
            "url": "/web/content/?model=%s&id=%s&field=result_file&filename_field=result_filename&download=true" % (self._name, self.id),
            "target": "self",
        }
