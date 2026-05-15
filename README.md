# Párovač bankovních plateb a pohledávek v Excelu

Skript v Pythonu pro automatické párování bankovních výpisů (CSV, standardní formát ABO/GPC) s interní evidencí neuhrazených faktur v Excelu. Provádí párování 1:1, identifikuje klientské odchylky (částečné platby, přeplatky, neexistující nebo chybějící variabilní symboly), aktualizuje stavy v Excelu a generuje přehled dlužníků po splatnosti včetně podkladů pro upomínky.

## Motivace a řešený problém

V kancelářské praxi je běžné, že účetní nebo asistentka denně stahuje z banky výpis plateb a v Excelu ručně dohledává podle variabilního symbolu a částky, kteří odběratelé již zaplatili a kdo je po splatnosti.

Tento proces je při větším počtu faktur zdlouhavý a náchylný k chybám – snadno dojde k přehlédnutí částečné úhrady nebo platby se špatně zadaným symbolem. Skript tuto rutinu automatizuje: během několika sekund zkontroluje všechny platby, barevně aktualizuje zdrojovou tabulku a připraví podklady pro další řešení.

## Jak skript funguje

1. **Načtení dat**: Skript načte interní evidenci pohledávek v Excelu (`.xlsx`) a bankovní výpis. Podporován je jak běžný formát CSV (s automatickou detekcí oddělovače středník nebo čárka), tak standardní český bankovní formát ABO/GPC (zpracování transakčních řádků typu 075).
2. **Párování 1:1**: Hledá se shoda mezi variabilním symbolem a částkou:
   - Pokud částka i variabilní symbol přesně sedí, faktura se označí stavem `Uhrazeno`, doplní se datum platby, uhrazená částka a řádek se podbarví zeleně.
3. **Klientské odchylky**:
   - **Částečná úhrada**: Klient poslal méně, než je hodnota faktury. Stav se nastaví na `Částečně uhrazeno`, zaznamená se přijatá částka a vypočte se nedoplatek.
   - **Přeplatek**: Klient poslal vyšší částku. Zaznamená se výše přeplatku pro účetní narovnání.
   - **Neznámý nebo chybějící VS**: Platby, které nelze podle symbolu přiřadit k žádné faktuře, se odloží na samostatný list k manuálnímu prověření (včetně čísla účtu a zprávy pro příjemce).
4. **Kontrola splatnosti a urgence**:
   - U všech neuhrazených faktur (včetně těch s částečnou platbou) skript porovná datum splatnosti s referenčním datem.
   - Pohledávky po splatnosti se zařadí do seznamu k urgenci s rozlišením podle délky prodlení:
     - 1 až 7 dní: 1. Přátelská upomínka
     - 8 až 30 dní: 2. Důrazná upomínka
     - více než 30 dní: 3. Předžalobní výzva
   - Pro každého dlužníka se vygeneruje text výzvy k úhradě s rekapitulací dlužné částky a dnů po splatnosti.

## Struktura výstupního Excel souboru

Výsledkem běhu skriptu je sešit formátu `.xlsx` se čtyřmi listy:

- **Souhrn**: Přehledná tabulka s celkovými čísly (celková suma pohledávek, spárovaná částka, zbývající dluh po splatnosti, počty spárovaných a problematických plateb).
- **Aktualizovane_pohledavky**: Původní seznam faktur s doplněným stavem, datem úhrady, uhrazenou částkou a barevným rozlišením (zelená = uhrazeno, žlutá = částečně uhrazeno, červená = po splatnosti).
- **Dluznici_k_urgenci**: Přehled všech neuhrazených faktur po splatnosti seřazený sestupně podle dnů prodlení, včetně navrženého textu upomínky.
- **Klientske_odchylky**: Seznam transakcí z banky, které nebylo možné automaticky spárovat 1:1 (nedoplatky, přeplatky, neznámé nebo chybějící VS).

## Použité technologie

- **Python 3.10+**
- **pandas**: Načtení, filtrování a transformace tabulkových dat.
- **openpyxl**: Tvorba, formátování a stylování výstupního Excel souboru (šířky sloupců, barevné výplně, ohraničení).
- **unittest**: Testovací sada pro ověření párovací logiky a okrajových případů.

## Požadavky a instalace

Pro běh je potřeba Python 3.10 nebo novější. Závislosti se nainstalují příkazem:

```bash
pip install -r requirements.txt
```

Soubor `requirements.txt` obsahuje dvě knihovny:
- `pandas>=2.0.0`
- `openpyxl>=3.1.0`

## Použití

### Základní spuštění

Skript se spouští z příkazové řádky:

```bash
python main.py
```

Ve výchozím nastavení skript načte data ze složky `data/` a vygeneruje report do složky `output/`:
- Vstupní faktury: `data/pohledavky.xlsx`
- Vstupní výpis: `data/vypis_plateb.csv`
- Výstupní report: `output/vysledek_parovani.xlsx`

V terminálu se po doběhnutí vypíše textová rekapitulace s přehledem dlužníků a odchylek.

### Parametry příkazové řádky

Cesty k souborům i referenční datum pro výpočet splatnosti lze předat přes argumenty:

```bash
python main.py --excel cesta/k/fakturam.xlsx --bank cesta/k/vypisu.csv --output vystup.xlsx --date 2026-11-01
```

Přehled parametrů:
- `--excel`, `-e`: Cesta k Excel souboru s pohledávkami (výchozí: `data/pohledavky.xlsx`)
- `--bank`, `-b`: Cesta k bankovnímu výpisu CSV nebo GPC (výchozí: `data/vypis_plateb.csv`)
- `--output`, `-o`: Cesta k výstupnímu Excel reportu (výchozí: `output/vysledek_parovani.xlsx`)
- `--date`, `-d`: Referenční datum pro výpočet splatnosti ve formátu YYYY-MM-DD (výchozí: 2026-11-01)

## Testování a ukázková data

Projekt má oddělená uživatelská data od testovacích fixtures:

- `data/`: Složka určená pro vaše reálná data při běžném provozu.
- `test-data/`: Složka s fixními testovacími daty, která slouží pro automatické testy. Změny v `data/` nijak neovlivní výsledek testů.

Spuštění testů:

```bash
python -m unittest discover tests
```

Přegenerování vzorových dat:

```bash
python scripts/generate_sample_data.py
```

## Struktura repozitáře

```text
fakturace_skript/
├── main.py                     # Vstupní skript pro spuštění z CLI
├── src/
│   ├── __init__.py
│   └── reconciler.py           # Třída BankReconciler s párovací logikou a exportem
├── data/                       # Uživatelská data pro běžný běh
│   ├── pohledavky.xlsx
│   └── vypis_plateb.csv
├── test-data/                  # Izolovaná testovací data pro unittest
│   ├── pohledavky.xlsx
│   └── vypis_plateb.csv
├── scripts/
│   └── generate_sample_data.py # Pomocný skript pro vytvoření testovacích dat
├── tests/
│   └── test_reconciler.py      # Jednotkové testy
├── requirements.txt            # Seznam závislostí
├── .gitignore
└── README.md
```

## Formulace do životopisu / portfolia

**Párovač bankovních výpisů vůči pohledávkám (Python)**
- Vytvořil jsem skript pro automatické párování bankovních transakcí (CSV/GPC) vůči interní evidenci pohledávek v Excelu.
- Skript provádí párování 1:1 na základě variabilního symbolu a částky, zachycuje klientské odchylky (částečné úhrady, přeplatky, chybné nebo chybějící VS) a automaticky generuje podklady pro upomínání dlužníků podle délky prodlení.
- Technologie: Python, pandas, openpyxl, unittest.
