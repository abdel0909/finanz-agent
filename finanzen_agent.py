
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import argparse, datetime as dt, re, shutil, sys
from pathlib import Path
import pandas as pd

BASE = Path(__file__).parent.resolve()
DATA_DIR = BASE / "data"
INBOX_DIR, PROCESSED_DIR, REPORT_DIR = DATA_DIR/"inbox", DATA_DIR/"processed", BASE/"reports"
RULES_CSV, WORKBOOK = BASE/"rules.csv", BASE/"Budget.xlsx"

TX_COLS = ["date","account_id","payee","amount_eur","currency","category","tags","note","external_id","source"]

def ensure_dirs(): 
    for d in [INBOX_DIR, PROCESSED_DIR, REPORT_DIR]: d.mkdir(parents=True, exist_ok=True)

def load_rules():
    if not RULES_CSV.exists(): return pd.DataFrame(columns=["pattern","category","tags"])
    df = pd.read_csv(RULES_CSV); return df.fillna("")

def _read_csv_robust(path): 
    try: return pd.read_csv(path, sep=None, engine="python")
    except: return pd.read_csv(path)

def normalize_columns(df):
    colmap = {}
    for c in df.columns:
        cl = c.lower().strip()
        if cl in ["datum","date","booking date","buchungstag"]: colmap[c]="date"
        elif cl in ["verwendungszweck","empfänger","payee","beschreibung"]: colmap[c]="payee"
        elif cl in ["betrag","amount","amount_eur","value"]: colmap[c]="amount_eur"
        elif cl in ["währung","currency"]: colmap[c]="currency"
        elif cl in ["notiz","note"]: colmap[c]="note"
        elif cl in ["id","external_id","transaction id"]: colmap[c]="external_id"
        elif cl in ["konto","account","account_id"]: colmap[c]="account_id"
    df=df.rename(columns=colmap)
    for c in TX_COLS: 
        if c not in df: df[c]="" if c!="amount_eur" else 0.0
    df["date"]=pd.to_datetime(df["date"],errors="coerce").dt.date
    df["amount_eur"]=pd.to_numeric(df["amount_eur"],errors="coerce").fillna(0.0)
    return df[TX_COLS]

def read_new_csv_files():
    files=sorted(INBOX_DIR.glob("*.csv")); frames=[]
    for f in files: frames.append(normalize_columns(_read_csv_robust(f)))
    return (pd.concat(frames,ignore_index=True) if frames else pd.DataFrame(columns=TX_COLS)), files

def apply_rules(df,rules):
    mask=df["category"].astype(str).str.strip()==""
    for _,r in rules.iterrows():
        m=mask & df["payee"].str.upper().str.contains(str(r["pattern"]),regex=True,na=False)
        df.loc[m,"category"]=r["category"]
        if "tags" in r: df.loc[m,"tags"]=r["tags"]
    df.loc[df["category"].str.strip()=="","category"]="Sonstiges"
    return df

def write_workbook(tx):
    tx["year_month"]=pd.to_datetime(tx["date"]).dt.to_period("M").astype(str)
    monthly=tx.groupby(["year_month","category"],as_index=False)["amount_eur"].sum()
    with pd.ExcelWriter(WORKBOOK,engine="openpyxl",mode="a",if_sheet_exists="replace") as w:
        tx.to_excel(w,"Transactions",index=False); monthly.to_excel(w,"Monthly_Summary",index=False)

def generate_reports(tx):
    tx["year_month"]=pd.to_datetime(tx["date"]).dt.to_period("M").astype(str)
    for ym,g in tx.groupby("year_month"): g.to_csv(REPORT_DIR/f"report_{ym}.csv",index=False)

def move_processed(files): 
    [shutil.move(str(f),PROCESSED_DIR/f.name) for f in files]

def dedupe(tx): 
    key=tx["date"].astype(str)+"|"+tx["account_id"]+"|"+tx["amount_eur"].astype(str)+"|"+tx["payee"]+"|"+tx["external_id"]
    tx["__k"]=key; return tx.drop_duplicates("__k").drop(columns="__k")

def run_agent():
    ensure_dirs(); rules=load_rules()
    existing=pd.read_excel(WORKBOOK,"Transactions") if WORKBOOK.exists() else pd.DataFrame(columns=TX_COLS)
    new,files=read_new_csv_files()
    if new.empty: print("[INFO] keine neuen CSVs"); return
    new=apply_rules(new,rules); all_tx=dedupe(pd.concat([existing,new],ignore_index=True))
    write_workbook(all_tx); generate_reports(all_tx); move_processed(files)
    print("[DONE] aktualisiert")

if __name__=="__main__": run_agent()
