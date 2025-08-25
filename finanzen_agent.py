
#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import argparse
import datetime as dt
from pathlib import Path
import shutil
import pandas as pd

# ───────── Pfade
BASE = Path(__file__).parent.resolve()
DATA_DIR = BASE / "data"
INBOX_DIR = DATA_DIR / "inbox"
PROCESSED_DIR = DATA_DIR / "processed"
REPORT_DIR = BASE / "reports"
RULES_CSV = BASE / "rules.csv"
WORKBOOK = BASE / "Budget.xlsx"

# ───────── Schema
TX_COLS = ["date","account_id","payee","amount_eur","currency",
           "category","tags","note","external_id","source"]

DEFAULT_ACCOUNTS = pd.DataFrame({
    "account_id": ["Giro_DE1234","Kreditkarte_Visa","Bar","Trading_OANDA"],
    "bank_name":  ["Deutsche Bank","Visa","—","OANDA"],
    "owner":      ["","","",""],
    "note":       ["Hauptkonto","Online-Zahlungen","Bargeld","Trading-Konto"]
})
DEFAULT_CATEGORIES = pd.DataFrame({
    "category": [
        "Miete","Strom","Internet","Versicherung",
        "Lebensmittel","Drogerie & Haushalt","ÖPNV & Tanken","Gesundheit & Medikamente",
        "Freizeit & Reisen","Bekleidung","Kinder & Schule","Sonstiges",
        "Einnahmen: Gehalt","Einnahmen: Nebenjob","Einnahmen: Trading","Sparen & Bauspar",
        "Verpflegung","Beitrag Sport","Beitrag Spa","ChatGPT","Colab","TradingView"
    ],
    "budget_monthly_eur": [900,120,60,90,450,60,120,80,120,60,100,100,0,0,0,200,150,40,50,20,10,15]
})

# ───────── Helpers
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
    balances = pd.DataFrame(columns=["account_id","balance_eur","last_tx_date"])
    with pd.ExcelWriter(WORKBOOK, engine="openpyxl") as w:
        DEFAULT_ACCOUNTS.to_excel(w, "Accounts", index=False)
        DEFAULT_CATEGORIES.to_excel(w, "Categories", index=False)
        transactions.to_excel(w, "Transactions", index=False)
        monthly.to_excel(w, "Monthly_Summary", index=False)
        balances.to_excel(w, "Balances", index=False)

def load_existing_transactions():
    if not WORKBOOK.exists():
        return pd.DataFrame(columns=TX_COLS)
    try:
        df = pd.read_excel(WORKBOOK, sheet_name="Transactions", engine="openpyxl")
        for c in TX_COLS:
            if c not in df.columns:
                df[c] = "" if c != "amount_eur" else 0.0
        return df[TX_COLS]
    except Exception as e:
        print(f"[WARN] Transactions nicht lesbar ({e}) – starte leer.")
        return pd.DataFrame(columns=TX_COLS)

def load_rules():
    if not RULES_CSV.exists():
        return pd.DataFrame(columns=["pattern","category","tags"])
    df = pd.read_csv(RULES_CSV)
    for c in ["pattern","category","tags"]:
        if c not in df.columns:
            df[c] = ""
    return df.fillna("")

def _read_csv_robust(path: Path) -> pd.DataFrame:
    try:
        return pd.read_csv(path, sep=None, engine="python")
    except Exception:
        return pd.read_csv(path)

def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    # Map typische Bankspalten -> Standard
    colmap = {}
    for c in df.columns:
        cl = str(c).strip().lower()
        if cl in ["datum","date","booking date","buchungstag","wertstellung","valuta"]:
            colmap[c] = "date"
        elif cl in ["verwendungszweck","empfänger","empfaenger","beschreibung","payee"]:
            colmap[c] = "payee"
        elif cl in ["betrag","amount","amount_eur","value","umsatz"]:
            colmap[c] = "amount_eur"
        elif cl in ["währung","waehrung","currency"]:
            colmap[c] = "currency"
        elif cl in ["konto","account","account_id","iban"]:
            colmap[c] = "account_id"
        elif cl in ["notiz","note","vermerk","zweck"]:
            colmap[c] = "note"
        elif cl in ["id","external_id","buchungsid","transaction id","buchungsnr"]:
            colmap[c] = "external_id"
        elif cl in ["quelle","source"]:
            colmap[c] = "source"
        elif cl in ["kategorie","category"]:
            colmap[c] = "category"
        elif cl in ["tags","tag"]:
            colmap[c] = "tags"
    df = df.rename(columns=colmap)

    # fehlende Spalten ergänzen
    for c in TX_COLS:
        if c not in df.columns:
            df[c] = "" if c != "amount_eur" else 0.0

    # Typen säubern
    df["date"] = pd.to_datetime(df["date"], errors="coerce").dt.date
    df["amount_eur"] = pd.to_numeric(df["amount_eur"], errors="coerce").fillna(0.0)
    for c in ["payee","category","tags","note","external_id","account_id","currency","source"]:
        df[c] = df[c].astype(str)

    # Defaults
    if (df["currency"].str.strip() == "").all():
        df["currency"] = "EUR"
    if (df["account_id"].str.strip() == "").all():
        df["account_id"] = "Giro_DE1234"
    if (df["source"].str.strip() == "").all():
        df["source"] = "csv"

    return df[TX_COLS]

def read_new_csv_files(inbox: Path):
    files = sorted(inbox.glob("*.csv"))
    if not files:
        return pd.DataFrame(columns=TX_COLS), []
    frames = []
    for f in files:
        try:
            df = _read_csv_robust(f)
            frames.append(normalize_columns(df))
            print(f"[READ] {f.name} → {len(df)} Zeilen")
        except Exception as e:
            print(f"[ERROR] {f.name} nicht lesbar: {e}")
    return (pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(columns=TX_COLS)), files

def apply_rules(df: pd.DataFrame, rules: pd.DataFrame) -> pd.DataFrame:
    if rules.empty:
        df.loc[df["category"].str.strip() == "", "category"] = "Sonstiges"
        return df
    mask_empty = df["category"].astype(str).str.strip().eq("")
    payee_upper = df["payee"].fillna("").str.upper()
    for _, r in rules.iterrows():
        pat = str(r.get("pattern", "") or "").strip()
        if not pat:
            continue
        cat = str(r.get("category", "Sonstiges") or "Sonstiges")
        tag = str(r.get("tags", "") or "")
        m = mask_empty & payee_upper.str.contains(pat, regex=True, na=False)
        if m.any():
            df.loc[m, "category"] = cat
            cur = df.loc[m, "tags"].fillna("")
            df.loc[m, "tags"] = (cur + ("," if (cur != "") & (tag != "") else "") + tag).str.strip(",")
    df.loc[df["category"].astype(str).str.strip() == "", "category"] = "Sonstiges"
    return df

# ───────── Bulletproof Dedupe
def dedupe(transactions: pd.DataFrame) -> pd.DataFrame:
    """Duplikate sicher entfernen – alle Schlüsselspalten strikt zu Strings."""
    t = transactions.copy()

    # Alle Schlüsselspalten sicher belegen
    if "date" not in t: t["date"] = ""
    if "account_id" not in t: t["account_id"] = ""
    if "payee" not in t: t["payee"] = ""
    if "external_id" not in t: t["external_id"] = ""
    if "amount_eur" not in t: t["amount_eur"] = 0.0

    # Harte Typ-Konvertierung
    t["date"] = pd.to_datetime(t["date"], errors="coerce").dt.strftime("%Y-%m-%d").fillna("").astype(str)
    t["account_id"] = t["account_id"].fillna("").astype(str)
    t["payee"] = t["payee"].fillna("").astype(str)
    t["external_id"] = t["external_id"].fillna("").astype(str)
    t["amount_eur"] = pd.to_numeric(t["amount_eur"], errors="coerce").fillna(0.0).round(2).astype(str)

    # Schlüssel bilden
    t["__k"] = t["date"] + "|" + t["account_id"] + "|" + t["amount_eur"] + "|" + t["payee"] + "|" + t["external_id"]
    t = t.drop_duplicates(subset="__k").drop(columns="__k")
    return t

def compute_balances(transactions: pd.DataFrame) -> pd.DataFrame:
    tx = transactions.copy()
    if tx.empty:
        return pd.DataFrame(columns=["account_id","balance_eur","last_tx_date"])
    bal = tx.groupby("account_id", as_index=False)["amount_eur"].sum().rename(columns={"amount_eur":"balance_eur"})
    last = tx.groupby("account_id", as_index=False)["date"].max().rename(columns={"date":"last_tx_date"})
    bal = bal.merge(last, on="account_id", how="left")
    total = pd.DataFrame([{"account_id":"✔ Gesamt","balance_eur":bal["balance_eur"].sum(),"last_tx_date":""}])
    return pd.concat([bal, total], ignore_index=True)

def write_workbook(transactions: pd.DataFrame):
    tx = transactions.copy()
    tx["year_month"] = pd.to_datetime(tx["date"]).dt.to_period("M").astype(str)
    monthly = tx.groupby(["year_month","category"], as_index=False)["amount_eur"].sum().rename(columns={"amount_eur":"sum_eur"})
    balances = compute_balances(transactions)

    if WORKBOOK.exists():
        with pd.ExcelWriter(WORKBOOK, engine="openpyxl", mode="a", if_sheet_exists="replace") as w:
            tx.to_excel(w, "Transactions", index=False)
            monthly.to_excel(w, "Monthly_Summary", index=False)
            balances.to_excel(w, "Balances", index=False)
    else:
        with pd.ExcelWriter(WORKBOOK, engine="openpyxl") as w:
            DEFAULT_ACCOUNTS.to_excel(w, "Accounts", index=False)
            DEFAULT_CATEGORIES.to_excel(w, "Categories", index=False)
            tx.to_excel(w, "Transactions", index=False)
            monthly.to_excel(w, "Monthly_Summary", index=False)
            balances.to_excel(w, "Balances", index=False)

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
        if target.exists():
            stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
            target = PROCESSED_DIR / f"{f.stem}_{stamp}{f.suffix}"
        shutil.move(str(f), str(target))

# ───────── Main
def run_agent(dry_run: bool = False):
    ensure_dirs()
    # CSVs einlesen
    new_df, files = read_new_csv_files(INBOX_DIR)
    if new_df.empty:
        print("[INFO] Keine neuen CSVs in data/inbox/")
        return

    # Bestand + Regeln
    existing = load_existing_transactions()
    rules = load_rules()
    new_df = apply_rules(new_df, rules)

    # Merge + Dedupe
    all_tx = pd.concat([existing, new_df], ignore_index=True)
    all_tx = dedupe(all_tx)

    if dry_run:
        print(f"[DRY] Neue Zeilen: {len(new_df)} | Gesamt nach Merge/Dedupe: {len(all_tx)}")
        return

    # Schreiben + Reports + Ablage
    write_workbook(all_tx)
    generate_reports(all_tx)
    move_processed(files)
    print("[DONE] Workbook aktualisiert, Reports erstellt, CSVs verschoben.")

def main():
    p = argparse.ArgumentParser(description="Finanzen-Agent")
    p.add_argument("--dry-run", action="store_true", help="nur lesen/prüfen, nichts schreiben/verschieben")
    args = p.parse_args()
    if not WORKBOOK.exists():
        create_workbook_if_missing()
    run_agent(dry_run=args.dry_run)

if __name__ == "__main__":
    main()
