#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
notify.py – Tages/Monats/Quartals/Jahres-Report per E-Mail (Gmail SMTP)
Unabhängig vom Agent: erzeugt/repariert Budget.xlsx, baut Diagramme, sendet Mail.
"""

import os, ssl, smtplib, re
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from pathlib import Path
from datetime import datetime, date, timedelta
from zoneinfo import ZoneInfo
from zipfile import BadZipFile

import pandas as pd
import matplotlib.pyplot as plt

# ------------------- Pfade & Umgebung -------------------
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

# ------------------- Excel / Daten -------------------
def _write_empty_workbook() -> None:
    """Legt eine minimale Budget.xlsx mit benötigten Sheets an."""
    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    tx = pd.DataFrame(columns=[
        "date","account_id","payee","amount_eur","currency",
        "category","tags","note","external_id","source"
    ])
    monthly = pd.DataFrame(columns=["year_month","category","sum_eur"])
    balances = pd.DataFrame(columns=["account_id","balance_eur","last_tx_date"])
    with pd.ExcelWriter(WORKBOOK, engine="openpyxl") as w:
        tx.to_excel(w, "Transactions", index=False)
        monthly.to_excel(w, "Monthly_Summary", index=False)
        balances.to_excel(w, "Balances", index=False)
    print("[INFO] Neue Budget.xlsx mit Minimalstruktur angelegt.")

def _ensure_workbook_ok() -> None:
    """Sicherstellen, dass Budget.xlsx existiert und lesbar ist."""
    if not WORKBOOK.exists():
        _write_empty_workbook(); return
    try:
        pd.read_excel(WORKBOOK, sheet_name="Transactions", engine="openpyxl")
    except BadZipFile:
        print("[WARN] Budget.xlsx beschädigt → neu anlegen …")
        WORKBOOK.unlink(missing_ok=True)
        _write_empty_workbook()

def load_tx() -> pd.DataFrame:
    _ensure_workbook_ok()
    tx = pd.read_excel(WORKBOOK, sheet_name="Transactions", engine="openpyxl")
    if tx.empty:
        return tx
    tx["date"] = pd.to_datetime(tx["date"], errors="coerce").dt.date
    tx["amount_eur"] = pd.to_numeric(tx["amount_eur"], errors="coerce").fillna(0.0)
    tx["category"] = tx.get("category", "").astype(str)
    return tx

# ------------------- Perioden/Reporting -------------------
def _period_range(period: str, ref: date) -> tuple[date, date, str]:
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
        s = date(ref.year-1, 1, 1); e = date(ref.year-1, 12, 31); label = str(s.year)
    else:
        raise ValueError("period must be daily|monthly|quarterly|yearly")
    return s, e, label

def _sum_spent(df: pd.DataFrame) -> float:
    return float((-df.loc[df["amount_eur"] < 0, "amount_eur"]).sum())

def build_report(period: str, today: date | None = None) -> dict:
    tx = load_tx()
    if today is None: today = datetime.now(TZ).date()
    s, e, label = _period_range(period, today)
    mask = (tx["date"] >= s) & (tx["date"] <= e) if not tx.empty else []
    df = tx.loc[mask].copy() if not tx.empty else pd.DataFrame(columns=tx.columns)

    total_spent = _sum_spent(df) if not df.empty else 0.0
    income = float(df.loc[df["amount_eur"] > 0, "amount_eur"].sum()) if not df.empty else 0.0
    net = float(df["amount_eur"].sum()) if not df.empty else 0.0

    # Kategorien
    cats = pd.DataFrame(columns=["category","spent"])
    images: list[Path] = []
    (REPORT_DIR/"imgs").mkdir(parents=True, exist_ok=True)

    if not df.empty:
        cat = df[df["amount_eur"] < 0].groupby("category", as_index=False)["amount_eur"].sum()
        cat["spent"] = -cat["amount_eur"]
        cats = cat[["category","spent"]].sort_values("spent", ascending=False)

        # Balken-Chart Top-Kategorien
        if not cats.empty:
            fig = plt.figure()
            top = cats.head(10).sort_values("spent")
            plt.barh(top["category"], top["spent"])
            plt.title(f"Top-Kategorien {period} {label}")
            plt.xlabel("EUR")
            p = REPORT_DIR/"imgs"/f"{period}_{label}_categories.png"
            fig.tight_layout(); fig.savefig(p); plt.close(fig)
            images.append(p)

        # Zeitreihe kumulierte Ausgaben
        if period in ("monthly","quarterly","yearly"):
            fig = plt.figure()
            day = df.groupby("date", as_index=False)["amount_eur"].sum()
            day["cum_spent"] = (-day["amount_eur"].clip(upper=0)).cumsum()
            plt.plot(day["date"], day["cum_spent"])
            plt.title(f"Kumulierte Ausgaben {period} {label}")
            plt.xlabel("Datum"); plt.ylabel("EUR")
            p = REPORT_DIR/"imgs"/f"{period}_{label}_timeseries.png"
            fig.tight_layout(); fig.savefig(p); plt.close(fig)
            images.append(p)

    # einfache Bewertung
    savings_rate = (income + net) / income if income > 0 else 0.0
    rating = "✅ gut"
    if income == 0 and total_spent > 0: rating = "⚠️ nur Ausgaben"
    elif savings_rate < 0.05:          rating = "⚠️ sehr niedrig"
    elif savings_rate < 0.15:          rating = "🟡 mittel"

    return {
        "period": period, "label": label, "start": s, "end": e,
        "spent": round(total_spent, 2), "income": round(income, 2), "net": round(net, 2),
        "top_categories": cats.to_dict("records"), "images": images, "rating": rating
    }

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
    <hr/>{imgs}
    """

# ------------------- Mailversand (robust) -------------------
def send_email(subject: str, html: str, attachments: list[Path]) -> None:
    if not (SMTP_USER and SMTP_PASS and EMAIL_FROM and EMAIL_TO):
        raise RuntimeError("SMTP/Gmail Variablen fehlen (SMTP_USER/SMTP_PASS/EMAIL_FROM/EMAIL_TO).")

    # Fallback-Body & Plain-Text
    if not html or not html.strip():
        html = "<p>(Kein HTML-Inhalt erzeugt)</p>"
    plain = re.sub(r"<[^>]+>", "", html).replace("&nbsp;", " ").strip() or "(Kein Inhalt)"

    # Debug: Body ablegen
    debug_dir = REPORT_DIR / "debug"
    debug_dir.mkdir(parents=True, exist_ok=True)
    (debug_dir / "last_mail.html").write_text(html, encoding="utf-8")
    (debug_dir / "last_mail.txt").write_text(plain, encoding="utf-8")

    # multipart/mixed (Anhänge) + multipart/alternative (plain+html)
    msg = MIMEMultipart("mixed")
    msg["Subject"] = subject
    msg["From"] = EMAIL_FROM
    msg["To"] = ",".join(EMAIL_TO)

    alt = MIMEMultipart("alternative")
    alt.attach(MIMEText(plain, "plain", "utf-8"))
    alt.attach(MIMEText(html,  "html",  "utf-8"))
    msg.attach(alt)

    for p in attachments:
        with open(p, "rb") as f:
            part = MIMEApplication(f.read(), Name=p.name)
        part["Content-Disposition"] = f'attachment; filename="{p.name}"'
        msg.attach(part)

    ctx = ssl.create_default_context()
    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as s:
        s.starttls(context=ctx)
        s.login(SMTP_USER, SMTP_PASS)
        s.sendmail(EMAIL_FROM, EMAIL_TO, msg.as_string())

# ------------------- Orchestrierung -------------------
def _daily_at_20_local() -> bool:
    return datetime.now(TZ).hour == 20

def run(period: str, send: bool = True) -> None:
    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    today = datetime.now(TZ).date()

    periods: list[str] = []
    if period == "auto":
        if _daily_at_20_local(): periods.append("daily")
        if today.day == 1:
            periods.append("monthly")
            if today.month in (1, 4, 7, 10): periods.append("quarterly")
            if today.month == 1: periods.append("yearly")
    else:
        periods = [period]

    for per in periods:
        rep = build_report(per, today=today)
        subject = f"Finanzen – {per} – {rep['label']} (Ausgaben {rep['spent']:.2f} €)"
        html = render_html(rep)
        atts = list(rep["images"])

        # kleine CSV-Beilagen
        if per == "daily":
            tx = load_tx()
            m = (tx["date"] >= rep["start"]) & (tx["date"] <= rep["end"]) if not tx.empty else []
            dayfile = REPORT_DIR / f"report_{rep['label']}_day.csv"
            (tx.loc[m] if not tx.empty else tx).to_csv(dayfile, index=False); atts.append(dayfile)
        elif per == "monthly":
            mon = REPORT_DIR / f"report_{rep['label']}.csv"
            if mon.exists(): atts.append(mon)

        print(f"[REPORT] {per} {rep['label']} | spent={rep['spent']:.2f} | atts={len(atts)}")
        print(f"[DEBUG] HTML length: {len(html)} | to={EMAIL_TO}")
        if send:
            send_email(subject, html, atts)
            print(f"[MAIL] gesendet an: {EMAIL_TO}")

# ------------------- CLI -------------------
if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="Finanz-Reports per E-Mail senden")
    ap.add_argument("period", choices=["daily","monthly","quarterly","yearly","auto"], help="Zeitraum")
    ap.add_argument("--send", action="store_true", help="E-Mail wirklich senden (sonst nur erzeugen)")
    args = ap.parse_args()
    run(args.period, send=args.send)
