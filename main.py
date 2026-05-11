#!/usr/bin/env python3
"""Hlavní spouštěcí skript pro párování bankovních výpisů."""

import argparse
from datetime import date
from pathlib import Path
from src.reconciler import BankReconciler


def main():
    parser = argparse.ArgumentParser(
        description="Automated Bank Reconciliation & Receivables Matching Tool"
    )
    parser.add_argument(
        "--excel",
        "-e",
        default="data/pohledavky.xlsx",
        help="Cesta k souboru s pohledávkami v Excelu (výchozí: data/pohledavky.xlsx)",
    )
    parser.add_argument(
        "--bank",
        "-b",
        default="data/vypis_plateb.csv",
        help="Cesta k bankovnímu výpisu CSV (výchozí: data/vypis_plateb.csv)",
    )
    parser.add_argument(
        "--output",
        "-o",
        default="output/vysledek_parovani.xlsx",
        help="Cesta k výstupnímu Excel reportu (výchozí: output/vysledek_parovani.xlsx)",
    )
    parser.add_argument(
        "--date",
        "-d",
        default="2026-11-01",
        help="Referenční datum pro výpočet splatnosti YYYY-MM-DD (výchozí: 2026-11-01)",
    )

    args = parser.parse_args()

    ref_date = date.fromisoformat(args.date) if args.date else date.today()

    print("=" * 70)
    print("   AUTOMATED BANK RECONCILIATION & RECEIVABLES MATCHING TOOL")
    print("=" * 70)
    print(f"[*] Kniha pohledávek : {args.excel}")
    print(f"[*] Bankovní výpis   : {args.bank}")
    print(f"[*] Referenční datum : {ref_date}")
    print(f"[*] Výstupní soubor  : {args.output}")
    print("-" * 70)

    reconciler = BankReconciler(
        excel_path=args.excel,
        bank_path=args.bank,
        today=ref_date,
    )

    df_inv, df_debtors, df_anomalies = reconciler.reconcile()
    output_path = reconciler.save_excel_report(args.output)

    # Statistika
    total_invoices = len(df_inv)
    paid_count = len(df_inv[df_inv["Stav"] == "Uhrazeno"])
    partial_count = len(df_inv[df_inv["Stav"] == "Částečně uhrazeno"])
    debtors_count = len(df_debtors)
    anomalies_count = len(df_anomalies)

    total_amount = df_inv["Castka"].sum()
    paid_amount = df_inv["Uhrazena_castka"].sum()
    unpaid_amount = df_inv["Zbyva_uhradit"].sum()

    print("\n[VÝSLEDKY PÁROVÁNÍ]")
    print(f"  Celkem faktur v knize       : {total_invoices} (v celkové výši {total_amount:,.2f} Kč)")
    print(f"  Plně uhrazeno (1:1)         : {paid_count}")
    print(f"  Částečně uhrazeno           : {partial_count}")
    print(f"  Celkem připsáno z plateb    : {paid_amount:,.2f} Kč")
    print(f"  Zbývá k úhradě              : {unpaid_amount:,.2f} Kč")
    print(f"  Faktur po splatnosti        : {debtors_count}")
    print(f"  Klientských odchylek        : {anomalies_count}")

    if not df_debtors.empty:
        print("\n[PŘEHLED DLUŽNÍKŮ K URGENCI]")
        for _, row in df_debtors.iterrows():
            print(
                f"  ! {row['Odberatel']:<28} | Fa: {row['Cislo_faktury']} | "
                f"Po splatnosti: {row['Dny_po_splatnosti']:>2} dní | "
                f"K úhradě: {row['Zbyva_uhradit']:>10,.2f} Kč | {row['Uroven_urgenci']}"
            )

    if not df_anomalies.empty:
        print("\n[KLIENTSKÉ ODCHYLKY K PROŠETŘENÍ]")
        for _, row in df_anomalies.iterrows():
            print(
                f"  ? TX: {row.get('id_transakce', 'N/A')} | "
                f"VS: {str(row.get('variabilni_symbol', '')) or 'CHYBÍ':<8} | "
                f"Částka: {float(row.get('castka', 0)):>10,.2f} Kč | "
                f"Důvod: {row.get('duvod', '')}"
            )

    print("\n" + "=" * 70)
    print(f"[HOTOVO] Aktualizovaný report úspěšně uložen do:\n  -> {output_path.resolve()}")
    print("=" * 70)


if __name__ == "__main__":
    main()
