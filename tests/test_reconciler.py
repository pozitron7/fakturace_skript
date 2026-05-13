"""Automatické testy pro BankReconciler."""

from datetime import date
from pathlib import Path
import unittest
import pandas as pd
from src.reconciler import BankReconciler


class TestBankReconciler(unittest.TestCase):
    def setUp(self):
        self.base_dir = Path(__file__).resolve().parent.parent
        self.excel_path = self.base_dir / "test-data" / "pohledavky.xlsx"
        self.bank_path = self.base_dir / "test-data" / "vypis_plateb.csv"
        self.today = date(2026, 11, 1)
        self.reconciler = BankReconciler(
            excel_path=self.excel_path,
            bank_path=self.bank_path,
            today=self.today,
        )

    def test_normalize_vs(self):
        self.assertEqual(self.reconciler._normalize_vs("002026001"), "2026001")
        self.assertEqual(self.reconciler._normalize_vs(2026001), "2026001")
        self.assertEqual(self.reconciler._normalize_vs("2026001.0"), "2026001")
        self.assertEqual(self.reconciler._normalize_vs(None), "")
        self.assertEqual(self.reconciler._normalize_vs("nan"), "")

    def test_reconciliation_matching(self):
        df_inv, df_debtors, df_anomalies = self.reconciler.reconcile()

        # 1. Plně spárované faktury 1:1
        paid_invoices = df_inv[df_inv["Stav"] == "Uhrazeno"]
        self.assertEqual(len(paid_invoices), 3)
        self.assertTrue(all(paid_invoices["Variabilni_symbol"].isin(["2026001", "2026002", "2026003"])))

        # 2. Částečná úhrada (VF-2026-004)
        partial_inv = df_inv[df_inv["Variabilni_symbol"] == "2026004"].iloc[0]
        self.assertEqual(partial_inv["Stav"], "Částečně uhrazeno")
        self.assertEqual(partial_inv["Uhrazena_castka"], 20000.0)
        self.assertEqual(partial_inv["Zbyva_uhradit"], 12000.0)

        # 3. Přeplatek (VF-2026-005)
        overpay_inv = df_inv[df_inv["Variabilni_symbol"] == "2026005"].iloc[0]
        self.assertEqual(overpay_inv["Stav"], "Přeplatek")
        self.assertEqual(overpay_inv["Uhrazena_castka"], 13000.0)

        # 4. Klientské odchylky
        # - částečná úhrada (2026004)
        # - přeplatek (2026005)
        # - neznámý VS (9999999)
        # - chybějící VS
        self.assertEqual(len(df_anomalies), 4)

        # 5. Dlužníci po splatnosti
        # Celkem 4: VF-2026-006 (36 dní), VF-2026-007 (18 dní), VF-2026-008 (5 dní), VF-2026-004 (4 dny)
        # VF-2026-009 není po splatnosti
        self.assertEqual(len(df_debtors), 4)
        zeta = df_debtors[df_debtors["Variabilni_symbol"] == "2026006"].iloc[0]
        self.assertEqual(zeta["Uroven_urgenci"], "3. Předžalobní výzva")


if __name__ == "__main__":
    unittest.main()
