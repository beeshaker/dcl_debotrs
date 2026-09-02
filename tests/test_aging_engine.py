from datetime import date
from odoo.tests.common import TransactionCase

from ..models.aging_engine import calculate_debtor_ageing


class TestAgingEngine(TransactionCase):
    def test_fifo_receipt_clears_oldest_debt(self):
        rows = [
            {"month": date(2026, 1, 1), "arrears_bf": 100.0, "rent_levy": 50.0, "recoveries": 0, "adjustments": 0, "receipts": 0, "current_bal": 150.0},
            {"month": date(2026, 2, 1), "arrears_bf": 150.0, "rent_levy": 50.0, "recoveries": 0, "adjustments": 0, "receipts": -120.0, "current_bal": 80.0},
        ]
        result = calculate_debtor_ageing("1", "Test Debtor", "Test Site", rows, 80.0)
        self.assertAlmostEqual(result.aged["Opening"], 0.0)
        self.assertAlmostEqual(result.aged["Jan 2026"], 30.0)
        self.assertAlmostEqual(result.aged["Feb 2026"], 50.0)
        self.assertFalse(result.exception)
