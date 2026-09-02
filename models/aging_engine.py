from collections import OrderedDict
from dataclasses import dataclass, field
from datetime import date, datetime
import re


MONTH_NAMES = {
    "jan": 1, "january": 1,
    "feb": 2, "february": 2,
    "mar": 3, "march": 3,
    "apr": 4, "april": 4,
    "may": 5,
    "jun": 6, "june": 6,
    "jul": 7, "july": 7,
    "aug": 8, "august": 8,
    "sep": 9, "sept": 9, "september": 9,
    "oct": 10, "october": 10,
    "nov": 11, "november": 11,
    "dec": 12, "december": 12,
}


def as_number(value):
    if value in (None, ""):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        text = value.strip().replace(",", "")
        if text.startswith("(") and text.endswith(")"):
            text = "-" + text[1:-1]
        try:
            return float(text)
        except ValueError:
            return 0.0
    return 0.0


def parse_month(value, fallback_year=None):
    if isinstance(value, datetime):
        return date(value.year, value.month, 1)
    if isinstance(value, date):
        return date(value.year, value.month, 1)
    if not value:
        return None
    text = re.sub(r"[^A-Za-z0-9 ]+", " ", str(value)).strip().lower()
    parts = text.split()
    month = None
    year = None
    for part in parts:
        key = part[:3] if part[:3] in MONTH_NAMES else part
        if key in MONTH_NAMES:
            month = MONTH_NAMES[key]
        elif part.isdigit() and len(part) == 4:
            year = int(part)
    if month and (year or fallback_year):
        return date(year or fallback_year, month, 1)
    return None


def normalize_header(value):
    if value is None:
        return ""
    text = str(value).strip().lower().replace("/", " ")
    text = re.sub(r"\s+", " ", text)
    aliases = {
        "arrears b f": "arrears_bf",
        "arrears b/f": "arrears_bf",
        "rent levy": "rent_levy",
        "recoveries": "recoveries",
        "adjustments": "adjustments",
        "receipts": "receipts",
        "current bal": "current_bal",
    }
    return aliases.get(text, text.replace(" ", "_"))


@dataclass
class DebtorAgeResult:
    debtor_id: str
    debtor_name: str
    property_name: str
    billing_frequency: str
    opening_balance: float
    period_charges: float
    period_receipts: float
    current_outstanding: float
    aged: OrderedDict = field(default_factory=OrderedDict)
    calculated_outstanding: float = 0.0
    exception: str = ""


def _apply_credit(balances, amount):
    """Apply a positive credit/payment against the oldest positive debt first."""
    remaining = max(float(amount), 0.0)
    for key in balances:
        if remaining <= 0:
            break
        available = max(balances[key], 0.0)
        used = min(available, remaining)
        balances[key] -= used
        remaining -= used
    return remaining


def _apply_charge(balances, bucket, amount, unapplied_credit):
    charge = max(float(amount), 0.0)
    if unapplied_credit > 0:
        used = min(charge, unapplied_credit)
        charge -= used
        unapplied_credit -= used
    if charge:
        balances[bucket] += charge
    return unapplied_credit


def infer_billing_frequency(monthly_rows):
    billed = 0
    for row in monthly_rows:
        if abs(row.get("rent_levy", 0.0)) > 0.005:
            billed += 1
    total = len(monthly_rows)
    if billed == 0:
        return "Irregular"
    if total >= 6 and billed >= max(5, total - 2):
        return "Monthly"
    if billed in (2, 3, 4):
        return "Quarterly"
    if billed == 1:
        return "Annual / Irregular"
    return "Irregular"


def calculate_debtor_ageing(debtor_id, debtor_name, property_name, monthly_rows, source_closing=None):
    """
    monthly_rows must be chronological and contain:
    month, arrears_bf, rent_levy, recoveries, adjustments, receipts, current_bal.

    Positive charges are originated in their month. Negative adjustments and receipts
    are applied oldest-first (FIFO). Excess credits are retained as unallocated credit
    and therefore do not appear as negative ageing buckets.
    """
    if not monthly_rows:
        raise ValueError("No monthly rows supplied")

    rows = sorted(monthly_rows, key=lambda r: r["month"])
    opening = as_number(rows[0].get("arrears_bf"))
    balances = OrderedDict()
    balances["Opening"] = max(opening, 0.0)
    unapplied_credit = max(-opening, 0.0)

    period_charges = 0.0
    period_receipts = 0.0

    for row in rows:
        label = row["month"].strftime("%b %Y")
        balances[label] = balances.get(label, 0.0)

        rent = as_number(row.get("rent_levy"))
        recoveries = as_number(row.get("recoveries"))
        adjustments = as_number(row.get("adjustments"))
        receipts = as_number(row.get("receipts"))

        # Charges are accumulated by origin month. Negative non-receipt movements act as credits.
        non_receipt = rent + recoveries + adjustments
        period_charges += non_receipt
        if non_receipt >= 0:
            unapplied_credit = _apply_charge(balances, label, non_receipt, unapplied_credit)
        else:
            unapplied_credit += _apply_credit(balances, -non_receipt)

        # Source ledgers normally store receipts as negative values. Positive receipt values
        # are treated as reversals/debits in the current month.
        period_receipts += receipts
        if receipts < 0:
            unapplied_credit += _apply_credit(balances, -receipts)
        elif receipts > 0:
            unapplied_credit = _apply_charge(balances, label, receipts, unapplied_credit)

    calculated = sum(max(v, 0.0) for v in balances.values()) - unapplied_credit
    source_closing = as_number(source_closing if source_closing is not None else rows[-1].get("current_bal"))
    exception = ""
    if abs(calculated - source_closing) > 0.05:
        exception = "Calculated ageing differs from source closing by %.2f" % (calculated - source_closing)

    return DebtorAgeResult(
        debtor_id=str(debtor_id or "").strip(),
        debtor_name=str(debtor_name or "").strip(),
        property_name=str(property_name or "").strip(),
        billing_frequency=infer_billing_frequency(rows),
        opening_balance=opening,
        period_charges=period_charges,
        period_receipts=period_receipts,
        current_outstanding=source_closing,
        aged=balances,
        calculated_outstanding=calculated,
        exception=exception,
    )
