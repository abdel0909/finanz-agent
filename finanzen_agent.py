
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Finanzen-Agent: Ein-/Ausgaben strukturiert verwalten
- Liest CSVs aus data/inbox/
- Auto-Kategorisierung per rules.csv (Regex auf PAYEE)
- Aktualisiert Budget.xlsx (Transactions + Monthly_Summary)
- Schreibt Monats-Reports in reports/
- Verschiebt verarbeitete CSVs nach data/processed/

Start lokal/GitHub Actions:
    pip install -r requirements.txt
    python finanzen_agent.py
"""

import argparse
import datetime as dt
import re
from pathlib import Path
import shutil
import sys

import pandas as pd

# --------------------- Pfade ---------------------
BASE = Path(__file__).parent.resolve()
DATA_DIR = BASE / "data"
INBOX_DIR = DATA_DIR / "inbox"
PROCESSED_DIR = DATA_DIR / "processed"
REPORT_DIR = BASE / "reports"
RULES_CSV = BASE / "rules.csv"
WORKBOOK = BASE / "Budget.xlsx"

# --------------------- Defaults ---------------------
DEFAULT_ACCOUNTS = pd.DataFrame({
    "account_id": ["Giro_DE1234", "Kreditkarte_Visa", "Bar", "Trading_OANDA"],
    "bank_name":  ["Deutsche Bank", "Visa", "—", "OANDA"],
    "owner":      ["Abdelkader Bouardja"]*4,
    "note":       ["Hauptkonto", "Online-Zahlungen", "Bargeld", "Trading-Konto"]
})

DEFAULT_CATEGORIES = pd.DataFrame({
    "category": [
        "Miete & Nebenkosten","Strom & Gas","Mobilfunk & Internet","Versicherungen",
        "Lebensmittel","Drogerie & Haushalt","ÖPNV & Tanken","Gesundheit & Medikamente",
        "Freizeit & Reisen","Bekleidung","Kinder & Schule","Sonstiges",
        "Einnahmen: Gehalt","Einnahmen: Nebenjob","Einnahmen: Trading","Sparen & Bauspar"
    ],
    "budget_monthly_eur": [900,120,60,90,450,60,120,80,120,60,100,100, 0,0,0,200]
})

TX_COLS = ["date","account_id","payee","amount_eur","currency",
           "category","tags","note","external_id","source"]


# --------------------- Helpers ---------------------
def ensure_dirs():
    INBOX_DIR.mkdir(parents=True, exist_ok=True)
    PROCESSED_DIR.mkdir(parents=True, exist_ok=True)
    REPORT_DIR.mkdir(parents=True, exist_ok=True)


def create_workbook_if_missing():
    if WORKBOOK.exists():
        return
    print(f"[INIT] Erzeuge {WORKBOOK.name} …")
    transactions = pd.DataFrame(columns=TX_COLS)
    monthly = pd.DataFrame(columns=["year_month","category","sum_eur"])
    with pd.ExcelWriter(WORKBOOK, engine="openpyxl") as w:
        DEFAULT_ACCOUNTS.to_excel(w, sheet_name="Accounts", index=False)
        DEFAULT_CATEGORIES.to_excel(w, sheet_name="Categories", index=False)
        transactions.to_excel(w, sheet_name="Transactions", index=False)
        monthly.to_excel(w, sheet_name="Monthly_Summary", index=False)


def load_existing_transactions():
    if not WORKBOOK.exists():
        return pd.DataFrame(columns=TX_COLS)
    try:
        df = pd.read_excel(WORKBOOK, sheet_name="Transactions")
        # Harmonisieren
        for c in TX_COLS:
            if c not in df.columns:
                df[c] = "" if c not in ["amount_eur"] else 0.0
        df = df[TX_COLS]
        return df
    except Exception as e:
        print(f"[WARN] Konnte Transactions aus {WORKBOOK.name} nicht lesen: {e}")
        return pd.DataFrame(columns=TX_COLS)


def load_rules():
    if not RULES_CSV.exists():
        print(f"[WARN] {RULES_CSV.name} nicht gefunden – alle unbekannten Buchungen → Sonstiges.")
        return pd.DataFrame(columns=["pattern","category","tags"])
    df = pd.read_csv(RULES_CSV)
    for col in ["pattern","category","tags"]:
        if col not in df.columns:
            df[col] = ""
    df["pattern"] = df["pattern"].astype(str)
    return df


def _read_csv_robust(path: Path) -> pd.DataFrame:
    # versucht Separator automatisch zu erkennen (z. B. ';' bei Bank-Exports)
    try:
        df = pd.read_csv(path, sep=None, engine="python")
    except Exception:
        df = pd.read_csv(path)
    return df


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    # Auto-Mapping typischer Exportspalten
    needed = set(TX_COLS)
    colmap = {}
    for c in df.columns:
        cl = c.strip().lower()
        if cl in ["datum","date","booking date","buchungstag","wertstellung","valuta"]:
            colmap[c] = "date"
        elif cl in ["verwendungszweck","empfänger","payee","beschreibung","empfaenger"]:
            colmap[c] = "payee"
        elif cl in ["betrag","amount","amount_eur","value","umsatz"]:
            colmap[c] = "amount_eur"
        elif cl in ["währung","waehrung","currency"]:
            colmap[c] = "currency"
        elif cl in ["notiz","note","zweck","vermerk"]:
            colmap[c] = "note"
        elif cl in ["id","external_id","buchungsid","transaction id","buchungsnr"]:
            colmap[c] = "external_id"
        elif cl in ["konto","account","account_id","iban"]:
            colmap[c] = "account_id"
    df = df.rename(columns=colmap)

    # fehlende Spalten ergänzen
    for col in TX_COLS:
        if col not in df.columns:
            df[col] = "" if col not in ["amount_eur"] else 0.0

    # Defaults
    if (df["currency"].astype(str).str.strip() == "").all():
        df["currency"] = "EUR"
    if (df["account_id"].astype(str).str.strip() == "").all():
        df["account_id"] = "Giro_DE1234"
    if (df["source"].astype(str).str.strip() == "").all():
        df["source"] = "csv"

    # Typen
    df["date"] = pd.to_datetime(df["date"], errors="coerce").dt.date
    df["amount_eur"] = pd.to_numeric(df["amount_eur"], errors="coerce").fillna(0.0)
    for c in ["payee","category","tags","note","external_id","account_id","currency","source"]:
        df[c] = df[c].astype(str)

    return df[TX_COLS]


def read_new_csv_files(inbox: Path):
    files = sorted(inbox.glob("*.csv"))
    if not files:
        return pd.DataFrame(columns=TX_COLS), []

    frames = []
    for f in files:
        try:
            df = _read_csv_robust(f)
            df = normalize_columns(df)
            frames.append(df)
            print(f"[READ] {f.name} → {len(df)} Zeilen")
        except Exception as e:
            print(f"[ERROR] {f.name} konnte nicht gelesen werden: {e}")
    if not frames:
        return pd.DataFrame(columns=TX_COLS), []
    return pd.concat(frames, ignore_index=True), files


def apply_rules(df: pd.DataFrame, rules: pd.DataFrame) -> pd.DataFrame:
    if rules.empty:
        df.loc[df["category"].str.strip() == "", "category"] = "Sonstiges"
        return df

    mask_empty = df["category"].astype(str).str.strip().eq("")
    payee_upper = df["payee"].fillna("").str.upper()

    for _, r in rules.iterrows():
        pat = str(r.get("pattern", ""))
        if not pat or pat.strip() == "":
            continue
        cat = str(r.get("category", "Sonstiges"))
        tag = str(r.get("tags", "") or "")
        m = mask_empty & payee_upper.str.contains(pat, regex=True, na=False)
        if m.any():
            df.loc[m, "category"] = cat
            # Tags anhängen (ohne Doppel-Komma)
            cur = df.loc[m, "tags"].fillna("")
            df.loc[m, "tags"] = (cur + ("," if (cur != "") else "") + tag).str.strip(",")
    df.loc[df["category"].astype(str).str.strip() == "", "category"] = "Sonstiges"
    return df


def write_workbook(transactions: pd.DataFrame):
    # Monthly Summary berechnen
    tx = transactions.copy()
    tx["year_month"] = pd.to_datetime(tx["date"]).dt.to_period("M").astype(str)
    monthly = tx.groupby(["year_month","category"], as_index=False)["amount_eur"].sum()
    monthly = monthly.rename(columns={"amount_eur":"sum_eur"})

    try:
        with pd.ExcelWriter(WORKBOOK, engine="openpyxl", mode="a", if_sheet_exists="replace") as w:
            transactions.to_excel(w, sheet_name="Transactions", index=False)
            monthly.to_excel(w, sheet_name="Monthly_Summary", index=False)
    except FileNotFoundError:
        with pd.ExcelWriter(WORKBOOK, engine="openpyxl") as w:
            DEFAULT_ACCOUNTS.to_excel(w, sheet_name="Accounts", index=False)
            DEFAULT_CATEGORIES.to_excel(w, sheet_name="Categories", index=False)
            transactions.to_excel(w, sheet_name="Transactions", index=False)
            monthly.to_excel(w, sheet_name="Monthly_Summary", index=False)


def generate_reports(transactions: pd.DataFrame):
    tx = transactions.copy()
    tx["year_month"] = pd.to_datetime(tx["date"]).dt.to_period("M").astype(str)
    for ym, g in tx.groupby("year_month"):
        out = REPORT_DIR / f"report_{ym}.csv"
        g.to_csv(out, index=False)
        print(f"[REPORT] {out.name} geschrieben")


def move_processed(files):
    for f in files:
        target = PROCESSED_DIR / f.name
        # bei Namenskonflikt suffix anhängen
        if target.exists():
            stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
            target = PROCESSED_DIR / f"{f.stem}_{stamp}{f.suffix}"
        shutil.move(str(f), str(target))


def dedupe(transactions: pd.DataFrame) -> pd.DataFrame:
    t = transactions.copy()
    key = (
        t["date"].astype(str)
        + "|" + t["account_id"].astype(str)
        + "|" + t["amount_eur"].astype(str)
        + "|" + t["payee"].astype(str)
        + "|" + t["external_id"].astype(str)
    )
    t["__k"] = key
    t = t.drop_duplicates(subset="__k").drop(columns="__k")
    return t


def run_agent(dry_run: bool = False):
    ensure_dirs()
    create_workbook_if_missing()

    existing = load_existing_transactions()
    rules = load_rules()
    new_df, files = read_new_csv_files(INBOX_DIR)

    if new_df.empty:
        print("[INFO] Keine neuen CSVs in data/inbox/")
        return

    new_df = apply_rules(new_df, rules)
    all_tx = pd.concat([existing, new_df], ignore_index=True)
    all_tx = dedupe(all_tx)

    if dry_run:
        print(f"[DRY] {len(new_df)} neue Zeilen, gesamt {len(all_tx)} nach Merge/Dedupe.")
        return

    write_workbook(all_tx)
    generate_reports(all_tx)
    move_processed(files)

    print("[DONE] Workbook aktualisiert, Reports erstellt, CSVs verschoben.")


# --------------------- CLI ---------------------
def main():
    p = argparse.ArgumentParser(description="Finanzen-Agent")
    p.add_argument("--dry-run", action="store_true", help="nur einlesen & berichten, nichts schreiben/verschieben")
    args = p.parse_args()
    run_agent(dry_run=args.dry_run)


if __name__ == "__main__":
    main()
