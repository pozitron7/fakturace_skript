"""Generátor ukázkových testovacích dat pro párovač bankovních výpisů.

Vytváří:
- test-data/pohledavky.xlsx a test-data/vypis_plateb.csv (fixní data pro automatické testy)
- data/pohledavky.xlsx a data/vypis_plateb.csv (výchozí složka pro běžné spouštění)
"""

import argparse
from datetime import date, timedelta
from pathlib import Path
import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side


def generate_sample_data(target_dirs: list[Path] | None = None):
    base_dir = Path(__file__).resolve().parent.parent

    if target_dirs is None:
        target_dirs = [base_dir / "test-data", base_dir / "data"]

    today = date(2026, 11, 1)

    # 1. Evidence pohledávek (Excel)
    pohledavky = [
        {
            "Cislo_faktury": "VF-2026-001",
            "Variabilni_symbol": "2026001",
            "Odberatel": "ALFA Consulting s.r.o.",
            "ICO": "12345678",
            "Datum_vystaveni": (today - timedelta(days=25)).strftime("%Y-%m-%d"),
            "Datum_splatnosti": (today - timedelta(days=11)).strftime("%Y-%m-%d"),
            "Castka": 15400.0,
            "Mena": "CZK",
            "Stav": "Neuhrazeno",
            "Datum_uhrady": "",
            "Uhrazena_castka": 0.0,
            "Poznamka": "",
        },
        {
            "Cislo_faktury": "VF-2026-002",
            "Variabilni_symbol": "2026002",
            "Odberatel": "BETA Solutions s.r.o.",
            "ICO": "23456789",
            "Datum_vystaveni": (today - timedelta(days=20)).strftime("%Y-%m-%d"),
            "Datum_splatnosti": (today - timedelta(days=6)).strftime("%Y-%m-%d"),
            "Castka": 48250.0,
            "Mena": "CZK",
            "Stav": "Neuhrazeno",
            "Datum_uhrady": "",
            "Uhrazena_castka": 0.0,
            "Poznamka": "",
        },
        {
            "Cislo_faktury": "VF-2026-003",
            "Variabilni_symbol": "2026003",
            "Odberatel": "GAMMA Trade a.s.",
            "ICO": "34567890",
            "Datum_vystaveni": (today - timedelta(days=15)).strftime("%Y-%m-%d"),
            "Datum_splatnosti": (today - timedelta(days=1)).strftime("%Y-%m-%d"),
            "Castka": 9800.0,
            "Mena": "CZK",
            "Stav": "Neuhrazeno",
            "Datum_uhrady": "",
            "Uhrazena_castka": 0.0,
            "Poznamka": "",
        },
        {
            "Cislo_faktury": "VF-2026-004",
            "Variabilni_symbol": "2026004",
            "Odberatel": "DELTA Logistics s.r.o.",
            "ICO": "45678901",
            "Datum_vystaveni": (today - timedelta(days=18)).strftime("%Y-%m-%d"),
            "Datum_splatnosti": (today - timedelta(days=4)).strftime("%Y-%m-%d"),
            "Castka": 32000.0,
            "Mena": "CZK",
            "Stav": "Neuhrazeno",
            "Datum_uhrady": "",
            "Uhrazena_castka": 0.0,
            "Poznamka": "",
        },
        {
            "Cislo_faktury": "VF-2026-005",
            "Variabilni_symbol": "2026005",
            "Odberatel": "EPSILON Tech s.r.o.",
            "ICO": "56789012",
            "Datum_vystaveni": (today - timedelta(days=14)).strftime("%Y-%m-%d"),
            "Datum_splatnosti": (today - timedelta(days=2)).strftime("%Y-%m-%d"),
            "Castka": 12500.0,
            "Mena": "CZK",
            "Stav": "Neuhrazeno",
            "Datum_uhrady": "",
            "Uhrazena_castka": 0.0,
            "Poznamka": "",
        },
        {
            "Cislo_faktury": "VF-2026-006",
            "Variabilni_symbol": "2026006",
            "Odberatel": "ZETA Retail s.r.o.",
            "ICO": "67890123",
            "Datum_vystaveni": (today - timedelta(days=50)).strftime("%Y-%m-%d"),
            "Datum_splatnosti": (today - timedelta(days=36)).strftime("%Y-%m-%d"),
            "Castka": 75000.0,
            "Mena": "CZK",
            "Stav": "Neuhrazeno",
            "Datum_uhrady": "",
            "Uhrazena_castka": 0.0,
            "Poznamka": "Klient nereaguje na neformální upozornění",
        },
        {
            "Cislo_faktury": "VF-2026-007",
            "Variabilni_symbol": "2026007",
            "Odberatel": "ETA Stavby s.r.o.",
            "ICO": "78901234",
            "Datum_vystaveni": (today - timedelta(days=32)).strftime("%Y-%m-%d"),
            "Datum_splatnosti": (today - timedelta(days=18)).strftime("%Y-%m-%d"),
            "Castka": 28900.0,
            "Mena": "CZK",
            "Stav": "Neuhrazeno",
            "Datum_uhrady": "",
            "Uhrazena_castka": 0.0,
            "Poznamka": "",
        },
        {
            "Cislo_faktury": "VF-2026-008",
            "Variabilni_symbol": "2026008",
            "Odberatel": "THETA Design studio s.r.o.",
            "ICO": "89012345",
            "Datum_vystaveni": (today - timedelta(days=19)).strftime("%Y-%m-%d"),
            "Datum_splatnosti": (today - timedelta(days=5)).strftime("%Y-%m-%d"),
            "Castka": 14200.0,
            "Mena": "CZK",
            "Stav": "Neuhrazeno",
            "Datum_uhrady": "",
            "Uhrazena_castka": 0.0,
            "Poznamka": "",
        },
        {
            "Cislo_faktury": "VF-2026-009",
            "Variabilni_symbol": "2026009",
            "Odberatel": "IOTA Services s.r.o.",
            "ICO": "90123456",
            "Datum_vystaveni": (today - timedelta(days=4)).strftime("%Y-%m-%d"),
            "Datum_splatnosti": (today + timedelta(days=10)).strftime("%Y-%m-%d"),
            "Castka": 63000.0,
            "Mena": "CZK",
            "Stav": "Neuhrazeno",
            "Datum_uhrady": "",
            "Uhrazena_castka": 0.0,
            "Poznamka": "Standardní splatnost v běhu",
        },
    ]

    # 2. Bankovní výpis (CSV)
    bank_statement = [
        {
            "id_transakce": "TX-20261028-001",
            "datum": "2026-10-28",
            "castka": 15400.0,
            "mena": "CZK",
            "variabilni_symbol": "2026001",
            "protiucet": "123456789/0100",
            "nazev_protiuctu": "ALFA Consulting s.r.o.",
            "zprava_pro_prijemce": "Uhrada fa VF-2026-001",
        },
        {
            "id_transakce": "TX-20261029-002",
            "datum": "2026-10-29",
            "castka": 48250.0,
            "mena": "CZK",
            "variabilni_symbol": "2026002",
            "protiucet": "987654321/0800",
            "nazev_protiuctu": "BETA Solutions s.r.o.",
            "zprava_pro_prijemce": "Faktura 2026002",
        },
        {
            "id_transakce": "TX-20261030-003",
            "datum": "2026-10-30",
            "castka": 9800.0,
            "mena": "CZK",
            "variabilni_symbol": "2026003",
            "protiucet": "555666777/2010",
            "nazev_protiuctu": "GAMMA Trade a.s.",
            "zprava_pro_prijemce": "VF 2026003",
        },
        {
            "id_transakce": "TX-20261031-004",
            "datum": "2026-10-31",
            "castka": 20000.0,
            "mena": "CZK",
            "variabilni_symbol": "2026004",
            "protiucet": "444333222/0300",
            "nazev_protiuctu": "DELTA Logistics s.r.o.",
            "zprava_pro_prijemce": "1. cast faktury VF-2026-004",
        },
        {
            "id_transakce": "TX-20261031-005",
            "datum": "2026-10-31",
            "castka": 13000.0,
            "mena": "CZK",
            "variabilni_symbol": "2026005",
            "protiucet": "111222333/0600",
            "nazev_protiuctu": "EPSILON Tech s.r.o.",
            "zprava_pro_prijemce": "Platba objednavky 2026005",
        },
        {
            "id_transakce": "TX-20261101-006",
            "datum": "2026-11-01",
            "castka": 11000.0,
            "mena": "CZK",
            "variabilni_symbol": "9999999",
            "protiucet": "777888999/5500",
            "nazev_protiuctu": "KAPPA Invest s.r.o.",
            "zprava_pro_prijemce": "Zaloha na sluzby",
        },
        {
            "id_transakce": "TX-20261101-007",
            "datum": "2026-11-01",
            "castka": 5500.0,
            "mena": "CZK",
            "variabilni_symbol": "",
            "protiucet": "333222111/0100",
            "nazev_protiuctu": "LAMBDA Partners v.o.s.",
            "zprava_pro_prijemce": "Uhrada konzultaci bez VS",
        },
    ]

    df_pohledavky = pd.DataFrame(pohledavky)
    df_bank = pd.DataFrame(bank_statement)

    for target_dir in target_dirs:
        target_dir.mkdir(parents=True, exist_ok=True)

        excel_path = target_dir / "pohledavky.xlsx"
        with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
            df_pohledavky.to_excel(writer, index=False, sheet_name="Pohledavky")
            ws = writer.sheets["Pohledavky"]

            header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
            header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
            thin_border = Border(
                left=Side(style="thin", color="D9D9D9"),
                right=Side(style="thin", color="D9D9D9"),
                top=Side(style="thin", color="D9D9D9"),
                bottom=Side(style="thin", color="D9D9D9"),
            )

            for col_idx, col in enumerate(ws.iter_cols(min_row=1, max_row=ws.max_row), start=1):
                cell = col[0]
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")

                max_len = max(len(str(c.value or "")) for c in col)
                col_letter = openpyxl.utils.get_column_letter(col_idx)
                ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

                for c in col[1:]:
                    c.border = thin_border
                    c.alignment = Alignment(vertical="center")

        print(f"[OK] Vytvořen soubor pohledávek: {excel_path}")

        csv_path = target_dir / "vypis_plateb.csv"
        df_bank.to_csv(csv_path, sep=";", index=False, encoding="utf-8-sig")
        print(f"[OK] Vytvořen bankovní výpis: {csv_path}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Generátor testovacích dat")
    parser.add_argument(
        "--target",
        "-t",
        help="Cílová složka pro vygenerování (např. test-data nebo data). Pokud není zadáno, vygeneruje se do obou.",
    )
    args = parser.parse_args()

    if args.target:
        generate_sample_data([Path(args.target)])
    else:
        generate_sample_data()
