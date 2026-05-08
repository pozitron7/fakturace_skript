"""Bank Reconciliation & Receivables Matching Tool

Automatické párování bankovních plateb (CSV / GPC) vůči knize pohledávek v Excelu.
- 1:1 párování na základě Variabilního symbolu a Částky
- Detekce klientských odchylek (částečná úhrada, přeplatek, chybný/neznámý VS)
- Aktualizace stavu v Excelu (podbarvení, datum a výše úhrady)
- Generování seznamu dlužníků pro urgence a přehledu odchylek
"""

from datetime import date, datetime
from pathlib import Path
import pandas as pd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


class BankReconciler:
    def __init__(self, excel_path: Path | str, bank_path: Path | str, today: date | None = None):
        self.excel_path = Path(excel_path)
        self.bank_path = Path(bank_path)
        self.today = today or date.today()

    def _normalize_vs(self, val) -> str:
        if pd.isna(val) or val is None or str(val).strip().lower() in ("", "nan", "none"):
            return ""
        s = str(val).strip().split(".")[0].lstrip("0")
        return s

    def _parse_gpc(self) -> pd.DataFrame:
        """Načte bankovní výpis ve standardním formátu ABO (GPC)."""
        records = []
        with open(self.bank_path, "r", encoding="cp1250", errors="replace") as f:
            for line in f:
                if line.startswith("075"):
                    amount_cents = int(line[48:60].strip() or 0)
                    amount = amount_cents / 100.0
                    vs = line[61:71].strip().lstrip("0")
                    raw_date = line[81:87].strip()
                    try:
                        tx_date = datetime.strptime(raw_date, "%d%m%y").strftime("%Y-%m-%d")
                    except ValueError:
                        tx_date = self.today.strftime("%Y-%m-%d")
                    records.append({
                        "id_transakce": line[35:48].strip(),
                        "datum": tx_date,
                        "castka": amount,
                        "mena": "CZK",
                        "variabilni_symbol": vs,
                        "protiucet": line[19:35].strip(),
                        "nazev_protiuctu": line[97:].strip() if len(line) > 97 else "",
                        "zprava_pro_prijemce": line[71:97].strip() if len(line) > 71 else "",
                    })
        return pd.DataFrame(records)

    def load_data(self) -> tuple[pd.DataFrame, pd.DataFrame]:
        df_pohledavky = pd.read_excel(self.excel_path, dtype=object)
        df_pohledavky["Castka"] = pd.to_numeric(df_pohledavky["Castka"], errors="coerce").fillna(0.0)
        df_pohledavky["Uhrazena_castka"] = pd.to_numeric(
            df_pohledavky.get("Uhrazena_castka", 0.0), errors="coerce"
        ).fillna(0.0)
        df_pohledavky["Datum_uhrady"] = df_pohledavky.get("Datum_uhrady", "").fillna("").astype(str)
        df_pohledavky["Poznamka"] = df_pohledavky.get("Poznamka", "").fillna("").astype(str)
        df_pohledavky["Stav"] = df_pohledavky.get("Stav", "Neuhrazeno").fillna("Neuhrazeno").astype(str)
        df_pohledavky["VS_norm"] = df_pohledavky["Variabilni_symbol"].apply(self._normalize_vs)

        # Načtení bankovního výpisu: GPC vs CSV
        if self.bank_path.suffix.lower() == ".gpc":
            df_bank = self._parse_gpc()
        else:
            with open(self.bank_path, "r", encoding="utf-8-sig", errors="replace") as f:
                first_line = f.readline()
            sep = ";" if ";" in first_line else ","
            df_bank = pd.read_csv(self.bank_path, sep=sep, dtype=object, encoding="utf-8-sig")

        df_bank["castka"] = pd.to_numeric(df_bank["castka"], errors="coerce").fillna(0.0)
        df_bank["variabilni_symbol"] = df_bank["variabilni_symbol"].fillna("").astype(str)
        df_bank["vs_norm"] = df_bank["variabilni_symbol"].apply(self._normalize_vs)
        return df_pohledavky, df_bank

    def reconcile(self) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
        df_inv, df_bank = self.load_data()
        df_bank["stav_parovani"] = "Nespárováno"
        df_bank["odchylka_duvod"] = ""
        anomalies = []

        # 1. Párování plateb z banky
        for idx_b, b_row in df_bank.iterrows():
            vs_b = b_row["vs_norm"]
            amount_b = float(b_row["castka"])
            date_b = str(b_row.get("datum", self.today.strftime("%Y-%m-%d")))

            if not vs_b:
                df_bank.at[idx_b, "stav_parovani"] = "Odchylka"
                df_bank.at[idx_b, "odchylka_duvod"] = "Chybějící variabilní symbol"
                anomalies.append({**b_row.to_dict(), "duvod": "Chybějící variabilní symbol"})
                continue

            matches = df_inv[df_inv["VS_norm"] == vs_b]
            if matches.empty:
                df_bank.at[idx_b, "stav_parovani"] = "Odchylka"
                df_bank.at[idx_b, "odchylka_duvod"] = f"VS {vs_b} nenalezen v evidenci"
                anomalies.append({**b_row.to_dict(), "duvod": f"VS {vs_b} neexistuje v knize faktur"})
                continue

            idx_inv = matches.index[0]
            inv_amount = float(df_inv.at[idx_inv, "Castka"])

            # 1:1 shoda částky a VS
            if abs(amount_b - inv_amount) < 0.01:
                df_bank.at[idx_b, "stav_parovani"] = "Spárováno 1:1"
                df_inv.at[idx_inv, "Stav"] = "Uhrazeno"
                df_inv.at[idx_inv, "Datum_uhrady"] = date_b
                df_inv.at[idx_inv, "Uhrazena_castka"] = amount_b
                df_inv.at[idx_inv, "Poznamka"] = "Úhrada spárována automaticky 1:1"
            elif amount_b < inv_amount:
                diff = inv_amount - amount_b
                df_bank.at[idx_b, "stav_parovani"] = "Odchylka"
                df_bank.at[idx_b, "odchylka_duvod"] = f"Částečná úhrada (nedoplatek {diff:,.2f} Kč)"
                df_inv.at[idx_inv, "Stav"] = "Částečně uhrazeno"
                df_inv.at[idx_inv, "Datum_uhrady"] = date_b
                df_inv.at[idx_inv, "Uhrazena_castka"] = amount_b
                df_inv.at[idx_inv, "Poznamka"] = f"Klientská odchylka: Nedoplatek {diff:,.2f} Kč"
                anomalies.append({**b_row.to_dict(), "duvod": f"Částečná úhrada: schází {diff:,.2f} Kč"})
            else:
                overpay = amount_b - inv_amount
                df_bank.at[idx_b, "stav_parovani"] = "Odchylka"
                df_bank.at[idx_b, "odchylka_duvod"] = f"Přeplatek ({overpay:,.2f} Kč)"
                df_inv.at[idx_inv, "Stav"] = "Přeplatek"
                df_inv.at[idx_inv, "Datum_uhrady"] = date_b
                df_inv.at[idx_inv, "Uhrazena_castka"] = amount_b
                df_inv.at[idx_inv, "Poznamka"] = f"Klientská odchylka: Přeplatek {overpay:,.2f} Kč"
                anomalies.append({**b_row.to_dict(), "duvod": f"Přeplatek o {overpay:,.2f} Kč"})

        # 2. Vyhodnocení neuhrazených faktur a dlužníků pro urgence
        df_inv["Datum_splatnosti_dt"] = pd.to_datetime(df_inv["Datum_splatnosti"]).dt.date
        df_inv["Dny_po_splatnosti"] = (self.today - df_inv["Datum_splatnosti_dt"]).apply(lambda x: max(0, x.days))
        df_inv["Zbyva_uhradit"] = df_inv["Castka"] - df_inv["Uhrazena_castka"]

        unpaid = df_inv[df_inv["Stav"] != "Uhrazeno"].copy()
        debtors = unpaid[(unpaid["Dny_po_splatnosti"] > 0) & (unpaid["Zbyva_uhradit"] > 0.01)].copy()

        def assign_urgency(days):
            if days <= 7:
                return "1. Přátelská upomínka"
            elif days <= 30:
                return "2. Důrazná upomínka"
            return "3. Předžalobní výzva"

        if not debtors.empty:
            debtors["Uroven_urgenci"] = debtors["Dny_po_splatnosti"].apply(assign_urgency)
            debtors["Navrh_textu_urgence"] = debtors.apply(
                lambda r: f"Urgence faktury {r['Cislo_faktury']} (VS {r['Variabilni_symbol']}): po splatnosti {r['Dny_po_splatnosti']} dní, zbývá uhradit {r['Zbyva_uhradit']:,.2f} Kč.",
                axis=1,
            )
            debtors = debtors.sort_values(by="Dny_po_splatnosti", ascending=False)
        else:
            debtors = pd.DataFrame()

        df_anomalies = pd.DataFrame(anomalies)
        if not df_anomalies.empty and "vs_norm" in df_anomalies.columns:
            df_anomalies = df_anomalies.drop(columns=["vs_norm"])

        df_inv = df_inv.drop(columns=["VS_norm", "Datum_splatnosti_dt"])
        return df_inv, debtors, df_anomalies

    def save_excel_report(self, output_path: Path | str) -> Path:
        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        df_inv, df_debtors, df_anomalies = self.reconcile()

        colors = {
            "header": PatternFill("solid", fgColor="1F497D"),
            "header_debt": PatternFill("solid", fgColor="C00000"),
            "header_warn": PatternFill("solid", fgColor="ED7D31"),
            "header_summary": PatternFill("solid", fgColor="366092"),
            "paid": PatternFill("solid", fgColor="E2EFDA"),
            "partial": PatternFill("solid", fgColor="FFF2CC"),
            "overdue": PatternFill("solid", fgColor="FCE4D6"),
        }
        font_hdr = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        border = Border(*(Side(style="thin", color="D9D9D9") for _ in range(4)))

        summary_data = [
            {"Ukazatel": "Celkový počet faktur v knize", "Hodnota": f"{len(df_inv)} ks"},
            {"Ukazatel": "Plně spárováno 1:1", "Hodnota": f"{len(df_inv[df_inv['Stav'] == 'Uhrazeno'])} ks"},
            {"Ukazatel": "Částečně uhrazeno", "Hodnota": f"{len(df_inv[df_inv['Stav'] == 'Částečně uhrazeno'])} ks"},
            {"Ukazatel": "Neuhrazeno po splatnosti (dlužníci)", "Hodnota": f"{len(df_debtors)} ks"},
            {"Ukazatel": "Klientských odchylek z banky", "Hodnota": f"{len(df_anomalies)} ks"},
            {"Ukazatel": "Celková hodnota pohledávek", "Hodnota": f"{df_inv['Castka'].sum():,.2f} Kč"},
            {"Ukazatel": "Celkem připsáno z výpisu", "Hodnota": f"{df_inv['Uhrazena_castka'].sum():,.2f} Kč"},
            {"Ukazatel": "Celkový dluh po splatnosti", "Hodnota": f"{df_debtors['Zbyva_uhradit'].sum():,.2f} Kč" if not df_debtors.empty else "0.00 Kč"},
        ]
        df_summary = pd.DataFrame(summary_data)

        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            df_summary.to_excel(writer, sheet_name="Souhrn", index=False)
            df_inv.to_excel(writer, sheet_name="Aktualizovane_pohledavky", index=False)
            if not df_debtors.empty:
                df_debtors.to_excel(writer, sheet_name="Dluznici_k_urgenci", index=False)
            if not df_anomalies.empty:
                df_anomalies.to_excel(writer, sheet_name="Klientske_odchylky", index=False)

            for name, sheet in writer.sheets.items():
                hdr_fill = colors["header_debt"] if "Dluznici" in name else (
                    colors["header_warn"] if "odchylky" in name else (
                        colors["header_summary"] if "Souhrn" in name else colors["header"]
                    )
                )
                for col_idx, col in enumerate(sheet.iter_cols(min_row=1, max_row=sheet.max_row), start=1):
                    col[0].fill, col[0].font, col[0].alignment = hdr_fill, font_hdr, Alignment(horizontal="center")
                    max_len = max(len(str(c.value or "")) for c in col)
                    sheet.column_dimensions[get_column_letter(col_idx)].width = max(max_len + 3, 14)
                    for c in col[1:]:
                        c.border = border
                        if name == "Aktualizovane_pohledavky" and col_idx == 8:  # Stav
                            c.fill = colors["paid"] if c.value == "Uhrazeno" else (colors["partial"] if "Částečně" in str(c.value) else colors["overdue"])

        return output_path
