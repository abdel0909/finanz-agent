#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
notify.py – PDF/CSV-Finanzberichte per E-Mail (Gmail SMTP)

- Liest Budget.xlsx (Transactions)
- Erzeugt Diagramm + kompaktes, farbiges PDF (Header, Kennzahlen, Tabelle, Diagramm)
- Versendet Daily/Monthly/Quarterly/Yearly oder automatisch um 20:00 lokal

Autor: Dein Finanz-Agent
"""

from __future__ import annotations
import os, ssl, smtplib, re
from pathlib import Path
from datetime import datetime, date, timedelta
from zoneinfo import ZoneInfo
from zipfile import BadZipFile
from typing import Tuple, List, Dict

import pandas as pd
import matplotlib.pyplot as plt

# ReportLab (PDF)
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, Flowable
)

# ------------------- Pfade & Umgebung -------------------
BASE        = Path(__file__).parent.resolve()
WORKBOOK    = BASE / "Budget.xlsx"
REPORT_DIR  = BASE / "reports"
IMG_DIR     = REPORT_DIR / "imgs"
DEBUG_DIR   = REPORT_DIR / "debug"
TZ          = ZoneInfo(os.getenv("LOCAL_TZ", "Europe/Berlin"))

SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.gmail.com")
SMTP_PORT   = int(os.getenv("SMTP_PORT", "587"))
SMTP_USER   = os.getenv("SMTP_USER", "bouardjaa@gmail.com")
SMTP_PASS   = os.getenv("SMTP_PASS", "zwqdwuyxdzydtaqu")
EMAIL_FROM  = os.getenv("bouardjaa@gmail.com", SMTP_USER or "")
EMAIL_TO    = [a.strip() for a in os.getenv("bouardjaa@gmail.com", SMTP_USER or "").split(",") if a.strip()]

PROFILE_NAME    = os.getenv("PROFILE_NAME", "").strip()
PROFILE_ADDRESS = os.getenv("PROFILE_ADDRESS", "").strip()

# ------------------- Excel-Absicherung -------------------
def _write_empty_workbook() -> None:
    """Legt minimale Budget.xlsx an."""
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
    """Stellt sicher, dass Budget.xlsx existiert und lesbar ist."""
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
    if "category" not in tx.columns:
        tx["category"] = ""
    tx["category"] = tx["category"].astype(str)
    return tx

# ------------------- Period & Kennzahlen -------------------
def _period_range(period: str, ref: date) -> Tuple[date, date, str]:
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
        label = f"{s.year}-Q{((s.month-1)//3)+1}"
    elif period == "yearly":
        s = date(ref.year-1, 1, 1); e = date(ref.year-1, 12, 31); label = str(s.year)
    else:
        raise ValueError("period must be daily|monthly|quarterly|yearly")
    return s, e, label

def _sum_spent(df: pd.DataFrame) -> float:
    return float((-df.loc[df["amount_eur"] < 0, "amount_eur"]).sum())

def build_report(period: str, today: date | None = None) -> Dict:
    tx = load_tx()
    if today is None: today = datetime.now(TZ).date()
    s, e, label = _period_range(period, today)

    if tx.empty:
        df = pd.DataFrame(columns=["date","account_id","payee","amount_eur","currency","category","tags","note"])
    else:
        mask = (tx["date"] >= s) & (tx["date"] <= e)
        df = tx.loc[mask, ["date","account_id","payee","amount_eur","currency","category","tags","note"]].copy()

    spent  = _sum_spent(df) if not df.empty else 0.0
    income = float(df.loc[df["amount_eur"] > 0, "amount_eur"].sum()) if not df.empty else 0.0
    net    = float(df["amount_eur"].sum()) if not df.empty else 0.0

    # Kategorien-Chart erzeugen (kompakt)
    IMG_DIR.mkdir(parents=True, exist_ok=True)
    images: List[Path] = []
    if not df.empty:
        cat = df[df["amount_eur"] < 0].groupby("category", as_index=False)["amount_eur"].sum()
        if not cat.empty:
            cat["spent"] = -cat["amount_eur"]
            top = cat.sort_values("spent", ascending=False).head(10).sort_values("spent")
            fig = plt.figure(figsize=(6, 3))  # kompakt
            plt.barh(top["category"], top["spent"])
            plt.title(f"Top-Kategorien {period} {label}")
            plt.xlabel("EUR")
            plt.tight_layout()
            p = IMG_DIR / f"{period}_{label}_categories.png"
            fig.savefig(p, dpi=150)
            plt.close(fig)
            images.append(p)

    # Bewertung (einfach)
    savings_rate = (income + net) / income if income > 0 else 0.0
    rating = "✅ gut"
    if income == 0 and spent > 0: rating = "⚠️ nur Ausgaben"
    elif savings_rate < 0.05:     rating = "⚠️ sehr niedrig"
    elif savings_rate < 0.15:     rating = "🟡 mittel"

    return {
        "period": period, "label": label, "start": s, "end": e,
        "spent": round(spent, 2), "income": round(income, 2), "net": round(net, 2),
        "df": df, "images": images, "rating": rating
    }

# ------------------- PDF Rendering -------------------
def _heading(elements, styles, rep):
    # Name/Adresse (wenn gesetzt)
    if PROFILE_NAME:
        elements.append(Paragraph(f"<b>{PROFILE_NAME}</b>", styles["Normal"]))
    if PROFILE_ADDRESS:
        elements.append(Paragraph(PROFILE_ADDRESS, styles["Normal"]))
    if PROFILE_NAME or PROFILE_ADDRESS:
        elements.append(Spacer(1, 6))

    # Titelblock
    elements.append(Paragraph("<para alignment='center'><font size=22><b>Finanz-Report</b></font></para>", styles["Normal"]))
    elements.append(Paragraph(
        f"<para alignment='center'><font size=12>{rep['period'].capitalize()} – {rep['label']}</font></para>",
        styles["Normal"])
    )
    elements.append(Paragraph(
        f"<para alignment='center'><font size=10>{rep['start']} bis {rep['end']}</font></para>",
        styles["Normal"])
    )
    elements.append(Spacer(1, 12))

def _cards(elements, rep):
    # Drei Kennzahlen-Karten in einer Zeile
    data = [
        ["Ausgaben",  f"{rep['spent']:.2f} €"],
        ["Einnahmen", f"{rep['income']:.2f} €"],
        ["Netto",     f"{rep['net']:.2f} €"],
    ]
    # Drei schmale Tabellen nebeneinander
    widths = [6.2*cm, 6.2*cm, 6.2*cm]
    row = []
    for i, (label, val) in enumerate(data):
        t = Table([[label],[f"<b>{val}</b>"]],
                  colWidths=[widths[i]],
                  rowHeights=[0.9*cm, 1.1*cm],
                  style=TableStyle([
                      ("BOX", (0,0), (-1,-1), 1, colors.black),
                      ("LINEABOVE", (0,0), (-1,0), 2, [colors.HexColor("#1E88E5"), colors.HexColor("#2E7D32"), colors.black][i]),
                      ("ALIGN", (0,1), (-1,1), "RIGHT"),
                      ("RIGHTPADDING", (0,1), (-1,1), 8),
                      ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
                      ("FONTSIZE", (0,0), (-1,-1), 10),
                  ]))
        row.append(t)
    elements.append(Table([row], colWidths=widths, style=TableStyle([])))
    elements.append(Spacer(1, 12))

def _tx_table(elements, df: pd.DataFrame, styles):
    elements.append(Paragraph("<b>Buchungen</b>", styles["Heading2"]))
    if df.empty:
        elements.append(Paragraph("(Keine Buchungen)", styles["Normal"]))
        elements.append(Spacer(1, 6))
        return

    show_cols = ["date","account_id","payee","amount_eur","currency","category","tags","note"]
    table_data = [show_cols] + df[show_cols].astype(str).values.tolist()

    col_widths = [2.2*cm, 2.6*cm, 3.2*cm, 2.2*cm, 1.6*cm, 2.8*cm, 3*cm, 4.5*cm]
    t = Table(table_data, repeatRows=1, colWidths=col_widths)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#1976D2")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("FONTSIZE", (0,0), (-1,0), 9),
        ("FONTSIZE", (0,1), (-1,-1), 8),
        ("ALIGN", (0,0), (-1,0), "CENTER"),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("GRID", (0,0), (-1,-1), 0.25, colors.black),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.whitesmoke, colors.Color(1,1,1)]),
    ]))
    elements.append(t)
    elements.append(Spacer(1, 8))
    ts = datetime.now(TZ).strftime("%Y-%m-%d %H:%M")
    elements.append(Paragraph(f"<font size=8>Erstellt am {ts}</font>", styles["Normal"]))
    elements.append(Spacer(1, 12))

def _chart(elements, rep):
    if not rep["images"]:
        return
    elements.append(Paragraph("<b>Diagramm</b>", getSampleStyleSheet()["Heading2"]))
    img_path = rep["images"][0]
    img = Image(str(img_path), width=14*cm, height=7*cm)  # kompakte Größe
    img.hAlign = "CENTER"
    elements.append(img)

def render_pdf_statement(rep: Dict, out: Path) -> Path:
    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    styles = getSampleStyleSheet()
    # etwas straffere Normal-Schrift
    styles["Normal"].fontName = "Helvetica"
    styles["Normal"].fontSize = 10

    elements: List[Flowable] = []
    _heading(elements, styles, rep)
    _cards(elements, rep)
    _tx_table(elements, rep["df"], styles)
    _chart(elements, rep)

    doc = SimpleDocTemplate(str(out), pagesize=A4,
                            leftMargin=1.5*cm, rightMargin=1.5*cm,
                            topMargin=1.5*cm, bottomMargin=1.5*cm)
    doc.build(elements)
    return out

# ------------------- E-Mail Versand -------------------
def send_email(subject: str, html: str, attachments: List[Path]) -> None:
    if not (SMTP_USER and SMTP_PASS and EMAIL_FROM and EMAIL_TO):
        raise RuntimeError("SMTP Variablen fehlen (SMTP_USER / SMTP_PASS / EMAIL_FROM / EMAIL_TO).")

    DEBUG_DIR.mkdir(parents=True, exist_ok=True)
    # Fallback Plain-Text
    if not html or not html.strip():
        html = "<p>(Kein HTML-Inhalt)</p>"
    plain = re.sub(r"<[^>]+>", "", html).strip() or "(Kein Inhalt)"

    (DEBUG_DIR / "last_mail.html").write_text(html, encoding="utf-8")
    (DEBUG_DIR / "last_mail.txt").write_text(plain, encoding="utf-8")

    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.application import MIMEApplication

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

# ------------------- HTML-Body (Kurz) -------------------
def render_html(rep: Dict) -> str:
    head = ""
    if PROFILE_NAME: head += f"<p><b>{PROFILE_NAME}</b></p>"
    if PROFILE_ADDRESS: head += f"<p>{PROFILE_ADDRESS}</p>"

    cats_ul = ""
    if not rep["df"].empty:
        cats = rep["df"].query("amount_eur<0").groupby("category")["amount_eur"].sum().sort_values()
        cats_ul = "".join(f"<li>{k}: {abs(v):.2f} €</li>" for k,v in cats.tail(5).items())
    cats_ul = cats_ul or "<li>(keine)</li>"

    return f"""
    {head}
    <h3>Finanz-Report – {rep['period']} – {rep['label']}</h3>
    <p><b>Zeitraum:</b> {rep['start']} bis {rep['end']}</p>
    <ul>
      <li><b>Ausgaben:</b> {rep['spent']:.2f} €</li>
      <li><b>Einnahmen:</b> {rep['income']:.2f} €</li>
      <li><b>Netto:</b> {rep['net']:.2f} €</li>
      <li><b>Bewertung:</b> {rep['rating']}</li>
    </ul>
    <p><b>Top-Kategorien (Ausgaben)</b></p>
    <ul>{cats_ul}</ul>
    """

# ------------------- Orchestrierung -------------------
def _daily_at_20_local() -> bool:
    return datetime.now(TZ).hour == 20

def run(period: str, send: bool = False) -> None:
    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    today = datetime.now(TZ).date()

    periods: List[str] = []
    if period == "auto":
        if _daily_at_20_local(): periods.append("daily")
        if today.day == 1:
            periods.append("monthly")
            if today.month in (1,4,7,10): periods.append("quarterly")
            if today.month == 1: periods.append("yearly")
    else:
        periods = [period]

    for per in periods:
        rep = build_report(per, today=today)

        # Dateien bauen
        pdf_path = REPORT_DIR / f"statement_{rep['label']}.pdf"
        render_pdf_statement(rep, pdf_path)

        # CSV beilegen
        csv_path = REPORT_DIR / f"report_{rep['label']}.csv"
        rep["df"].to_csv(csv_path, index=False)

        subject = f"Finanzen – {per} – {rep['label']} (Ausgaben {rep['spent']:.2f} €)"
        html = render_html(rep)
        atts = [pdf_path, csv_path]

        print(f"[REPORT] {per} {rep['label']} | spent={rep['spent']:.2f} | atts={len(atts)}")
        if send:
            send_email(subject, html, atts)
            print(f"[MAIL] gesendet an: {EMAIL_TO}")

# ------------------- CLI -------------------
if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="Finanz-Reports (PDF/CSV) per E-Mail senden")
    ap.add_argument("period", choices=["daily","monthly","quarterly","yearly","auto"], help="Zeitraum")
    ap.add_argument("--send", action="store_true", help="E-Mail wirklich senden (sonst nur Dateien erzeugen)")
    args = ap.parse_args()
    run(args.period, send=args.send)
