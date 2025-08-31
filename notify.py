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
# --- NEU: PDF-Renderer ersetzen --------------------------------------------
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, KeepTogether
from xml.sax.saxutils import escape as _esc

def _eur(x: float) -> str:
    try:
        s = f"{float(x):,.2f}"
    except Exception:
        return str(x)
    return s.replace(",", "X").replace(".", ",").replace("X", ".") + " €"

def _scaled_colwidths(target_width, mm_widths):
    pts = [w * mm for w in mm_widths]
    total = sum(pts)
    if total <= target_width:
        return pts
    f = target_width / total
    return [p * f for p in pts]

def _df_to_wrapped_table(df, doc_width):
    # gewünschte Spaltenreihenfolge (falls vorhanden)
    cols = [c for c in ["date","account_id","payee","amount_eur","currency","category","tags","note"] if c in df.columns]
    df = df[cols].copy()

    # Formatierungen
    if "amount_eur" in df.columns:
        df["amount_eur"] = df["amount_eur"].map(_eur)

    styles = getSampleStyleSheet()
    cell = ParagraphStyle(
        "cell", parent=styles["Normal"], fontSize=8, leading=10,
        spaceAfter=0, spaceBefore=0
    )
    # Header-Stil
    header = ParagraphStyle(
        "header", parent=styles["Normal"], fontSize=9, leading=11,
        textColor=colors.whitesmoke, spaceAfter=0, spaceBefore=0
    )

    # Head + Rows -> Paragraphs (wrapping)
    data = [[Paragraph(_esc(str(c)), header) for c in cols]]
    for _, r in df.iterrows():
        row = [Paragraph(_esc("" if pd.isna(r[c]) else str(r[c])), cell) for c in cols]
        data.append(row)

    # Spaltenbreiten (in mm, wird auf Seitenbreite skaliert)
    #             date  acct   payee  amount curr  cat    tags   note
    mm_widths = [ 22,   24,    50,    22,    16,   38,    38,    64 ]
    colWidths = _scaled_colwidths(doc_width, mm_widths[:len(cols)])

    tbl = Table(data, colWidths=colWidths, repeatRows=1, hAlign="LEFT")
    tbl.setStyle(TableStyle([
        # Header
        ("BACKGROUND", (0,0), (-1,0), colors.Color(0.10,0.33,0.85)),
        ("TEXTCOLOR",  (0,0), (-1,0), colors.whitesmoke),
        ("FONTNAME",   (0,0), (-1,0), "Helvetica-Bold"),
        ("ALIGN",      (0,0), (-1,0), "CENTER"),
        ("BOTTOMPADDING",(0,0),(-1,0), 6),

        # Body
        ("FONTNAME", (0,1), (-1,-1), "Helvetica"),
        ("FONTSIZE", (0,1), (-1,-1), 8),
        ("LEADING",  (0,1), (-1,-1), 10),
        ("VALIGN",   (0,1), (-1,-1), "TOP"),
        ("ROWBACKGROUNDS",(0,1),(-1,-1), [colors.whitesmoke, colors.Color(0.97,0.97,0.97)]),
        ("GRID", (0,0), (-1,-1), 0.3, colors.Color(0.80,0.80,0.85)),

        # etwas Luft
        ("LEFTPADDING", (0,0), (-1,-1), 4),
        ("RIGHTPADDING",(0,0), (-1,-1), 4),
        ("TOPPADDING",  (0,0), (-1,-1), 3),
        ("BOTTOMPADDING",(0,0),(-1,-1), 3),
    ]))
    return tbl

def render_pdf_statement(rep: dict, df: pd.DataFrame, images: list[Path], out_pdf: Path, profile: dict):
    """
    Erzeugt ein kompaktes, farbiges PDF mit:
    - Titel + Zeitraum
    - KPI-Leiste (Ausgaben/Einnahmen/Netto)
    - Diagramm direkt darunter (wenn vorhanden)
    - Buchungstabelle (schmale Spalten, Wrapping)
    """
    out_pdf.parent.mkdir(parents=True, exist_ok=True)

    # Dokument-Setup
    doc = SimpleDocTemplate(
        str(out_pdf),
        pagesize=A4,
        leftMargin=18*mm, rightMargin=18*mm,
        topMargin=16*mm, bottomMargin=16*mm
    )
    styles = getSampleStyleSheet()
    h1 = ParagraphStyle("h1", parent=styles["Heading1"], fontSize=24, leading=28, alignment=1)  # centered
    h2 = ParagraphStyle("h2", parent=styles["Heading2"], fontSize=12, leading=14, alignment=1)
    hsec = ParagraphStyle("hsec", parent=styles["Heading2"], fontSize=16, leading=19, spaceBefore=10, spaceAfter=6)
    small = ParagraphStyle("small", parent=styles["Normal"], fontSize=8, textColor=colors.grey)

    elems = []

    # Titel
    elems.append(Paragraph("Finanz-Report", h1))
    subtitle = f"{rep['period'].capitalize()} – {rep['label']}<br/>{rep['start']} bis {rep['end']}"
    elems.append(Paragraph(subtitle, h2))
    elems.append(Spacer(1, 6*mm))

    # KPI-Leiste (3 Boxen)
    k_label = ParagraphStyle("k_label", parent=styles["Normal"], fontSize=10, textColor=colors.grey)
    k_val   = ParagraphStyle("k_val",   parent=styles["Heading2"], fontSize=14, leading=16, textColor=colors.black)

    kpi_data = [
        [Paragraph("Ausgaben", k_label), Paragraph(f"<b>{_eur(rep['spent'])}</b>", k_val)],
        [Paragraph("Einnahmen", k_label), Paragraph(f"<b>{_eur(rep['income'])}</b>", k_val)],
        [Paragraph("Netto", k_label), Paragraph(f"<b>{_eur(rep['net'])}</b>", k_val)],
    ]

    # drei gleich breite Boxen über die Seitenbreite
    box_w = (doc.width - 2*mm) / 3.0
    kpi = Table(
        [[kpi_data[0]], [kpi_data[1]], [kpi_data[2]]],
        colWidths=[box_w], rowHeights=None, hAlign="LEFT"
    )
    kpi.setStyle(TableStyle([
        ("BOX", (0,0), (-1,-1), 0.9, colors.black),
        ("INNERGRID", (0,0), (-1,-1), 0.6, colors.black),
        ("LEFTPADDING",(0,0),(-1,-1), 6),
        ("RIGHTPADDING",(0,0),(-1,-1), 6),
        ("TOPPADDING",(0,0),(-1,-1), 6),
        ("BOTTOMPADDING",(0,0),(-1,-1), 8),
        # kleine Farbakzente oben
        ("LINEABOVE",(0,0),(0,0), 2, colors.Color(0.10,0.33,0.85)),  # blau
        ("LINEABOVE",(0,1),(0,1), 2, colors.Color(0.10,0.65,0.35)),  # grün
        ("LINEABOVE",(0,2),(0,2), 2, colors.black),
    ]))

    # kpi nebeneinander anordnen
    krow = Table([[kpi._cellvalues[0][0], kpi._cellvalues[1][0], kpi._cellvalues[2][0]]],
                 colWidths=[box_w, box_w, box_w], hAlign="LEFT", spaceBefore=0, spaceAfter=6)
    krow.setStyle(TableStyle([
        ("LEFTPADDING",(0,0),(-1,-1),0),
        ("RIGHTPADDING",(0,0),(-1,-1),0),
        ("TOPPADDING",(0,0),(-1,-1),0),
        ("BOTTOMPADDING",(0,0),(-1,-1),0),
    ]))
    elems.append(krow)
    elems.append(Spacer(1, 6*mm))

    # Diagramm direkt nach den KPIs
    if images:
        elems.append(Paragraph("Diagramm", hsec))
        img = Image(str(images[0]))
        # auf max 75% der Textbreite und 55 mm Höhe begrenzen (Seitenlayout)
        max_w = doc.width * 0.75
        max_h = 55 * mm
        img._restrictSize(max_w, max_h)
        elems.append(img)
        elems.append(Spacer(1, 6*mm))

    # Buchungen
    elems.append(Paragraph("Buchungen", hsec))
    if df is not None and not df.empty:
        table = _df_to_wrapped_table(df.copy(), doc.width)
        elems.append(table)
        elems.append(Spacer(1, 3*mm))
        elems.append(Paragraph(f"Erstellt am {datetime.now(TZ).strftime('%Y-%m-%d %H:%M')}", small))
    else:
        elems.append(Paragraph("<i>Keine Buchungen im Zeitraum.</i>", styles["Normal"]))

    doc.build(elems)
# --- Ende NEU ----------------------------------------------------------------

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
