#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os, ssl, smtplib, base64
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from pathlib import Path
from datetime import datetime, date, timedelta
from zoneinfo import ZoneInfo
from zipfile import BadZipFile

import pandas as pd
import matplotlib.pyplot as plt

# --- Pfade & Umgebung
BASE = Path(__file__).parent.resolve()
WORKBOOK = BASE / "Budget.xlsx"
REPORT_DIR = BASE / "reports"
TZ = ZoneInfo(os.getenv("LOCAL_TZ", "Europe/Berlin"))

SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.gmail.com")
SMTP_PORT   = int(os.getenv("SMTP_PORT", "587"))
SMTP_USER   = os.getenv("SMTP_USER", "bouardjaa@gmail.com")
SMTP_PASS   = os.getenv("SMTP_PASS", "zwqdwuyxdzydtaqu")
EMAIL_FROM  = os.getenv("bouardjaa@gmail.com", SMTP_USER)
EMAIL_TO    = [a.strip() for a in os.getenv("bouardjaa@gmail.com", SMTP_USER).split(",") if a.strip()]

# --- Helpers
def ensure_workbook():
    """Wenn Budget.xlsx fehlt/kaputt ist: minimale Struktur neu erstellen."""
    if not WORKBOOK.exists():
        _write_empty()
        return
    try:
        pd.read_excel(WORKBOOK, sheet_name="Transactions", engine="openpyxl")
    except BadZipFile:
        WORKBOOK.unlink(missing_ok=True)
        _write_empty()

def _write_empty():
    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    tx = pd.DataFrame(columns=["date","account_id","payee","amount_eur","currency",
                               "category","tags","note","external_id","source"])
    monthly = pd.DataFrame(columns=["year_month","category","sum_eur"])
    balances = pd.DataFrame(columns=["account_id","balance_eur","last_tx_date"])
    with pd.ExcelWriter(WORKBOOK, engine="openpyxl") as w:
        tx.to_excel(w, "Transactions", index=False)
        monthly.to_excel(w, "Monthly_Summary", index=False)
        balances.to_excel(w, "Balances", index=False)
    print("[INFO] Neue Budget.xlsx mit Minimalstruktur angelegt.")

def load_tx() -> pd.DataFrame:
    ensure_workbook()
    tx = pd.read_excel(WORKBOOK, sheet_name="Transactions", engine="openpyxl")
    if tx.empty:
        return tx
    tx["date"] = pd.to_datetime(tx["date"], errors="coerce").dt.date
    tx["amount_eur"] = pd.to_numeric(tx["amount_eur"], errors="coerce").fillna(0.0)
    tx["category"] = tx.get("category", "").astype(str)
    return tx

def period_range(period: str, ref: date):
    if period == "daily":
        s = e = ref; label = ref.strftime("%Y-%m-%d")
    elif period == "monthly":
        first_this = ref.replace(day=1)
        last_prev = first_this - timedelta(days=1)
        s = last_prev.replace(day=1); e = last_prev; label = s.strftime("%Y-%m")
    elif period == "quarterly":
        m = ((ref.month - 1)//3)*3 + 1
        first_this_q = date(ref.year, m, 1)
        last_prev = first_this_q - timedelta(days=1)
        qm = ((last_prev.month - 1)//3)*3 + 1
        s = date(last_prev.year, qm, 1)
        e = date(last_prev.year, qm+3, 1) - timedelta(days=1)
        label = f"{s.strftime('%Y-Q')}{(s.month-1)//3 + 1}"
    elif period == "yearly":
        s = date(ref.year-1,1,1); e = date(ref.year-1,12,31); label = str(s.year)
    else:
        raise ValueError("period must be daily|monthly|quarterly|yearly")
    return s,e,label

def build_report(period: str, today: date|None=None) -> dict:
    tx = load_tx()
    if today is None: today = datetime.now(TZ).date()
    s,e,label = period_range(period, today)
    mask = (tx["date"] >= s) & (tx["date"] <= e) if not tx.empty else []
    df = tx.loc[mask].copy() if not tx.empty else pd.DataFrame(columns=tx.columns if not tx.empty else [])

    spent = float((-df.loc[df["amount_eur"]<0,"amount_eur"]).sum()) if not df.empty else 0.0
    income = float(df.loc[df["amount_eur"]>0,"amount_eur"].sum()) if not df.empty else 0.0
    net = float(df["amount_eur"].sum()) if not df.empty else 0.0

    # Kategorien (nur Ausgaben)
    cats = pd.DataFrame(columns=["category","spent"])
    if not df.empty:
        cat = df[df["amount_eur"]<0].groupby("category",as_index=False)["amount_eur"].sum()
        cat["spent"] = -cat["amount_eur"]
        cats = cat[["category","spent"]].sort_values("spent",ascending=False)

    images=[]
    (REPORT_DIR/"imgs").mkdir(parents=True,exist_ok=True)
    if not cats.empty:
        fig = plt.figure()
        top = cats.tail(10).sort_values("spent")
        plt.barh(top["category"], top["spent"])
        plt.title(f"Top-Kategorien {period} {label}"); plt.xlabel("EUR")
        p = REPORT_DIR/"imgs"/f"{period}_{label}_categories.png"
        fig.tight_layout(); fig.savefig(p); plt.close(fig)
        images.append(p)

    if period in ("monthly","quarterly","yearly") and not df.empty:
        fig = plt.figure()
        day = df.groupby("date",as_index=False)["amount_eur"].sum()
        day["cum_spent"] = (-day["amount_eur"].clip(upper=0)).cumsum()
        plt.plot(day["date"], day["cum_spent"])
        plt.title(f"Kumulierte Ausgaben {period} {label}")
        plt.xlabel("Datum"); plt.ylabel("EUR")
        p = REPORT_DIR/"imgs"/f"{period}_{label}_timeseries.png"
        fig.tight_layout(); fig.savefig(p); plt.close(fig)
        images.append(p)

    # simple Bewertung
    savings_rate = (income + net)/income if income>0 else 0.0
    rating = "✅ gut"
    if income==0 and spent>0: rating="⚠️ nur Ausgaben"
    elif savings_rate < 0.05: rating="⚠️ sehr niedrig"
    elif savings_rate < 0.15: rating="🟡 mittel"

    return {"period":period,"label":label,"start":s,"end":e,
            "spent":round(spent,2),"income":round(income,2),"net":round(net,2),
            "top_categories":cats.to_dict("records"),"images":images,"rating":rating}

def render_html(rep: dict) -> str:
    cats = "".join([f"<li>{c['category']}: {c['spent']:.2f} €</li>" for c in rep["top_categories"][:10]]) or "<li>(keine)</li>"
    imgs = "".join([f'<p><b>{p.name}</b> (Anhang)</p>' for p in rep["images"]])
    return f"""
    <h2>Finanz-Report ({rep['period']} – {rep['label']})</h2>
    <p><b>Zeitraum:</b> {rep['start']} bis {rep['end']}</p>
    <ul>
      <li><b>Ausgaben</b>: {rep['spent']:.2f} €</li>
      <li><b>Einnahmen</b>: {rep['income']:.2f} €</li>
      <li><b>Netto</b>: {rep['net']:.2f} €</li>
      <li><b>Bewertung</b>: {rep['rating']}</li>
    </ul>
    <p><b>Top-Kategorien</b></p>
    <ul>{cats}</ul>
    <hr />{imgs}
    """

# Debug: Body speichern, damit du sehen kannst, was verschickt wird
(REPORT_DIR / "debug").mkdir(parents=True, exist_ok=True)
debug_html = REPORT_DIR / "debug" / "last_mail.html"
debug_txt  = REPORT_DIR / "debug" / "last_mail.txt"
debug_html.write_text(html, encoding="utf-8")
debug_txt.write_text(re.sub(r"<[^>]+>", "", html), encoding="utf-8")

print("[DEBUG] Subject:", subject)
print("[DEBUG] HTML length:", len(html))
print("[DEBUG] Attachments:", [a.name for a in attachments])

def maybe_send_daily_at_20() -> bool:
    return datetime.now(TZ).hour == 20

def run(period: str, send: bool=True):
    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    today = datetime.now(TZ).date()

    # Auto-Modus: täglich 20:00; am 1. zusätzlich Monat/Quartal/Jahr
    periods=[]
    if period=="auto":
        if maybe_send_daily_at_20(): periods.append("daily")
        if today.day==1:
            periods.append("monthly")
            if today.month in (1,4,7,10): periods.append("quarterly")
            if today.month==1: periods.append("yearly")
    else:
        periods=[period]

    for per in periods:
        rep = build_report(per, today=today)
        subject = f"Finanzen – {per} – {rep['label']} (Ausgaben {rep['spent']:.2f} €)"
        html = render_html(rep)
        atts = list(rep["images"])

        # kleine CSV-Beilagen
        if per=="daily":
            tx = load_tx()
            mask = (tx["date"]>=rep["start"]) & (tx["date"]<=rep["end"])
            dayfile = REPORT_DIR / f"report_{rep['label']}_day.csv"
            tx.loc[mask].to_csv(dayfile,index=False); atts.append(dayfile)
        elif per=="monthly":
            mon = REPORT_DIR / f"report_{rep['label']}.csv"
            if mon.exists(): atts.append(mon)

        print(f"[REPORT] {per} {rep['label']} | spent={rep['spent']:.2f} | atts={len(atts)}")
        if send:
            send_email(subject, html, atts)
            print(f"[MAIL] gesendet an: {EMAIL_TO}")

if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument("period", choices=["daily","monthly","quarterly","yearly","auto"])
    ap.add_argument("--send", action="store_true")
    args = ap.parse_args()
    run(args.period, send=args.send)
