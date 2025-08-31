#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
notify.py – Reports (Daily/Monthly/Quarterly/Yearly) als farbige PDF + Mail (Gmail)
- Liest Budget.xlsx (Sheet 'Transactions')
- baut Kennzahlen & Diagramm
- erzeugt PDF mit Kopf (Name/Adresse/Datum), KPIs, Tabelle (schmale Spalten), Diagramm
- verschickt die PDF + CSV-Anhang via Gmail SMTP
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

# --- ReportLab (PDF) ---
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
)

# ====================== Pfade & Umgebungen ======================
BASE        = Path(__file__).parent.resolve()
WORKBOOK    = BASE / "Budget.xlsx"
REPORT_DIR  = BASE / "reports"
IMG_DIR     = REPORT_DIR / "imgs"
DEBUG_DIR   = REPORT_DIR / "debug"
TZ          = ZoneInfo(os.getenv("LOCAL_TZ", "Europe/Berlin"))

# Absender/Empfänger (Gmail SMTP)
SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.gmail.com")
SMTP_PORT   = int(os.getenv("SMTP_PORT", "587"))
SMTP_USER   = os.getenv("SMTP_USER", "")
SMTP_PASS   = os.getenv("SMTP_PASS", "")
EMAIL_FROM  = os.getenv("EMAIL_FROM", SMTP_USER)
EMAIL_TO    = [a.strip() for a in os.getenv("EMAIL_TO", SMTP_USER).split(",") if a.strip()]

# Profil für PDF-Kopf
PROFILE = {
    "name":    os.getenv("PROFILE_NAME", "Dein Name"),
    "address": os.getenv("PROFILE_ADDRESS", "Straße 1, 12345 Stadt"),
}

# Farben / Styles
BLUE      = colors.HexColor("#1e66ff")
GREEN     = colors.HexColor("#2fbf71")
DARK      = colors.HexColor("#111111")
MUTED     = colors.HexColor("#666666")
LIGHT_BG  = colors.whitesmoke

styles = getSampleStyleSheet()
styles.add(ParagraphStyle(name="H1", fontSize=22, leading=26, alignment=1, spaceAfter=8, textColor=DARK))
styles.add(ParagraphStyle(name="H2", fontSize=12, leading=16, alignment=1, textColor=MUTED))
styles.add(ParagraphStyle(name="KPI", fontSize=16, leading=20, alignment=1, textColor=DARK))
styles.add(ParagraphStyle(name="KPI_Label", fontSize=10, leading=12, alignment=0, textColor=MUTED))
styles.add(ParagraphStyle(name="Section", fontSize=16, leading=20, textColor=DARK, spaceBefore=8, spaceAfter=6))
styles.add(ParagraphStyle(name="Small", fontSize=9, leading=11, textColor=MUTED))

# ====================== Excel laden / absichern ======================
def _write_empty_workbook() -> None:
    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    tx = pd.DataFrame(columns=[
        "date","account_id","payee","amount_eur","currency",
        "category","tags","note","external_id","source"
    ])
    with pd.ExcelWriter(WORKBOOK, engine="openpyxl") as w:
        tx.to_excel(w, "Transactions", index=False)
        pd.DataFrame(columns=["year_month","category","sum_eur"]).to_excel(w, "Monthly_Summary", index=False)
        pd.DataFrame(columns=["account_id","balance_eur","last_tx_date"]).to_excel(w, "Balances", index=False)
    print("[INFO] Neue Budget.xlsx mit Minimalstruktur angelegt.")

def _ensure_workbook_ok() -> None:
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
    tx["tags"] = tx.get("tags", "").astype(str)
    tx["note"] = tx.get("note", "").astype(str)
    return tx

# ====================== Periodik / Kennzahlen ======================
def _period_range(period: str, ref: date) -> tuple[date, date, str]:
    if period == "daily":
        s = e = ref; label = ref.strftime("%Y-%m-%d")
    elif period == "monthly":
        first_this = ref.replace(day=1)
        last_prev  = first_this - timedelta(days=1)
        s = last_prev.replace(day=1); e = last_prev; label = s.strftime("%Y-%m")
    elif period == "quarterly":
        m = ((ref.month - 1)//3)*3 + 1
        first_this_q = date(ref.year, m, 1)
        last_prev  = first_this_q - timedelta(days=1)
        qm = ((last_prev.month - 1)//3)*3 + 1
        s = date(last_prev.year, qm, 1)
        e = date(last_prev.year, qm+3, 1) - timedelta(days=1)
        label = f"{s.year}-Q{((s.month-1)//3)+1}"
    elif period == "yearly":
        s = date(ref.year-1, 1, 1); e = date(ref.year-1, 12, 31); label = str(s.year)
    else:
        raise ValueError("period must be daily|monthly|quarterly|yearly")
    return s, e, label

def build_report(period: str, today: date | None = None) -> dict:
    tx = load_tx()
    if today is None:
        today = datetime.now(TZ).date()
    s, e, label = _period_range(period, today)
    if tx.empty:
        df = tx.copy()
    else:
        df = tx.loc[(tx["date"] >= s) & (tx["date"] <= e)].copy()

    spent  = float((-df.loc[df["amount_eur"] < 0, "amount_eur"]).sum()) if not df.empty else 0.0
    income = float(df.loc[df["amount_eur"] > 0, "amount_eur"].sum()) if not df.empty else 0.0
    net    = float(df["amount_eur"].sum()) if not df.empty else 0.0

    # Top-Kategorien für Ausgaben
    cats_df = pd.DataFrame(columns=["category","spent"])
    if not df.empty:
        cat = df[df["amount_eur"] < 0].groupby("category", as_index=False)["amount_eur"].sum()
        cat["spent"] = -cat["amount_eur"]
        cats_df = cat[["category","spent"]].sort_values("spent", ascending=False)

    # Diagramm (schmal)
    IMG_DIR.mkdir(parents=True, exist_ok=True)
    chart = None
    if not cats_df.empty:
        top = cats_df.head(8).sort_values("spent")
        fig = plt.figure(figsize=(6.0, 2.3))  # schlank
        plt.barh(top["category"], top["spent"])
        plt.title(f"Top-Kategorien {period} {label}")
        plt.xlabel("EUR")
        chart = IMG_DIR / f"{period}_{label}_categories.png"
        fig.tight_layout(); fig.savefig(chart, dpi=160); plt.close(fig)

    # Bewertung
    savings_rate = (income + net) / income if income > 0 else 0.0
    rating = "✅ gut"
    if income == 0 and spent > 0: rating = "⚠️ nur Ausgaben"
    elif savings_rate < 0.05:     rating = "⚠️ sehr niedrig"
    elif savings_rate < 0.15:     rating = "🟡 mittel"

    return {
        "period": period, "label": label, "start": s, "end": e,
        "spent": round(spent, 2), "income": round(income, 2), "net": round(net, 2),
        "top_categories": cats_df.to_dict("records"),
        "chart": chart,
        "rating": rating,
        "rows": df.reset_index(drop=True)
    }

# ====================== PDF Rendering ======================
def _kpi_box(label: str, value: float, border_color=colors.black):
    box = Table(
        [
            [Paragraph(label, styles["KPI_Label"])],
            [Paragraph(f"<b>{value:,.2f} €</b>".replace(",", "X").replace(".", ",").replace("X", "."), styles["KPI"])],
        ],
        colWidths=[65*mm],
        rowHeights=[10*mm, 16*mm],
        style=TableStyle([
            ("BOX", (0,0), (-1,-1), 1, border_color),
            ("TOPPADDING", (0,0), (-1,-1), 3),
            ("BOTTOMPADDING", (0,0), (-1,-1), 3),
        ])
    )
    return box

def render_pdf_statement(rep: dict, out_pdf: Path, profile: dict) -> Path:
    REPORT_DIR.mkdir(parents=True, exist_ok=True)

    doc = SimpleDocTemplate(
        str(out_pdf),
        pagesize=A4,
        leftMargin=18*mm, rightMargin=18*mm, topMargin=18*mm, bottomMargin=16*mm,
        title=f"Finanz-Report {rep['period']} {rep['label']}"
    )

    flow = []

    # Kopf: Name/Adresse links, Datum rechts
    header_table = Table(
        [
            [
                Paragraph(f"<b>{profile.get('name','')}</b><br/>{profile.get('address','')}", styles["Small"]),
                Paragraph(datetime.now(TZ).strftime("%Y-%m-%d"), styles["Small"])
            ]
        ],
        colWidths=[120*mm, 40*mm],
        style=TableStyle([
            ("ALIGN", (1,0), (1,0), "RIGHT"),
        ])
    )
    flow += [header_table, Spacer(1, 6)]

    # Titel
    flow += [
        Paragraph("Finanz-Report", styles["H1"]),
        Paragraph(f"{rep['period'].capitalize()} – {rep['label']}", styles["H2"]),
        Paragraph(f"{rep['start']} bis {rep['end']}", styles["H2"]),
        Spacer(1, 6)
    ]

    # KPI-Reihe (3 Boxen)
    kpis = Table(
        [[
            _kpi_box("Ausgaben",  -rep["spent"], border_color=BLUE),   # negativer Betrag anzeigen? -> KPI zeigt als Betrag, daher -spent in Box
            _kpi_box("Einnahmen", rep["income"], border_color=GREEN),
            _kpi_box("Netto",     rep["net"],    border_color=colors.black),
        ]],
        colWidths=[65*mm, 65*mm, 65*mm],
        style=TableStyle([("VALIGN", (0,0), (-1,-1), "MIDDLE")])
    )
    flow += [kpis, Spacer(1, 8)]

    # Abschnitt: Buchungen
    flow += [Paragraph("Buchungen", styles["Section"])]

    df = rep["rows"].copy()
    if df.empty:
        flow += [Paragraph("Keine Buchungen im Zeitraum.", styles["Small"])]
    else:
        # Nur die wichtigsten Spalten & schmale Breiten
        view_cols = ["date","account_id","payee","amount_eur","currency","category","tags","note"]
        df = df[view_cols].fillna("")
        # Zahlen hübsch
        df["amount_eur"] = df["amount_eur"].map(lambda x: f"{x:,.2f} €".replace(",", "X").replace(".", ",").replace("X", "."))
        # Kopf + Daten
        data = [list(df.columns)] + df.values.tolist()

        # Spaltenbreiten (Summe < Seitenbreite)
        col_w_mm = [22, 26, 35, 24, 16, 28, 28, 40]  # in mm; insgesamt 219 mm -> passt mit Rändern (A4: 210mm; Innen: ~174mm), daher kleiner wählen:
        scale = (doc.width / mm) / sum(col_w_mm)
        col_w = [w*scale*mm for w in col_w_mm]

        tbl = Table(
            data,
            colWidths=col_w,
            repeatRows=1,
            style=TableStyle([
                ("BACKGROUND", (0,0), (-1,0), BLUE),
                ("TEXTCOLOR", (0,0), (-1,0), colors.white),
                ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
                ("ALIGN", (3,1), (3,-1), "RIGHT"),
                ("GRID", (0,0), (-1,-1), 0.3, colors.grey),
                ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.whitesmoke, colors.white]),
                ("FONTSIZE", (0,0), (-1,-1), 8),
                ("LEADING", (0,0), (-1,-1), 10),
            ])
        )
        flow += [tbl]

    flow += [Spacer(1, 6), Paragraph(f"Erstellt am {datetime.now(TZ).strftime('%Y-%m-%d %H:%M')}", styles["Small"]), Spacer(1, 8)]

    # Diagramm direkt danach (falls vorhanden)
    if rep["chart"] and Path(rep["chart"]).exists():
        max_w = doc.width * 0.75
        img = RLImage(str(rep["chart"]), width=max_w, height=55*mm)  # begrenzen
        flow += [Paragraph("Diagramm", styles["Section"]), img]

    doc.build(flow)
    return out_pdf

# ====================== Mailversand ======================
def send_email(subject: str, html: str, attachments: list[Path]) -> None:
    if not (SMTP_USER and SMTP_PASS and EMAIL_FROM and EMAIL_TO):
        raise RuntimeError("SMTP Variablen fehlen (SMTP_USER/SMTP_PASS/EMAIL_FROM/EMAIL_TO).")

    # Fallback Body
    if not html or not html.strip():
        html = "<p>(Kein Inhalt)</p>"
    plain = re.sub(r"<[^>]+>", "", html).strip() or "(Kein Inhalt)"

    DEBUG_DIR.mkdir(parents=True, exist_ok=True)
    (DEBUG_DIR / "last_mail.html").write_text(html, encoding="utf-8")
    (DEBUG_DIR / "last_mail.txt").write_text(plain, encoding="utf-8")

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

# ====================== Orchestrierung ======================
def _daily_at_20_local() -> bool:
    return datetime.now(TZ).hour == 20

def render_html_summary(rep: dict) -> str:
    return f"""
    <h3>Finanz-Report ({rep['period']} – {rep['label']})</h3>
    <p><b>Zeitraum:</b> {rep['start']} bis {rep['end']}</p>
    <ul>
      <li><b>Ausgaben:</b> {rep['spent']:.2f} €</li>
      <li><b>Einnahmen:</b> {rep['income']:.2f} €</li>
      <li><b>Netto:</b> {rep['net']:.2f} €</li>
      <li><b>Bewertung:</b> {rep['rating']}</li>
    </ul>
    """

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

    tx_all = load_tx()

    for per in periods:
        rep = build_report(per, today=today)
        subject = f"Finanzen – {per} – {rep['label']} (Ausgaben {rep['spent']:.2f} €)"

        # PDF
        pdf_file = REPORT_DIR / f"statement_{rep['label']}.pdf"
        render_pdf_statement(rep, pdf_file, PROFILE)

        # CSV-Tagesanhang (falls daily)
        atts = [pdf_file]
        if per == "daily":
            mask = (tx_all["date"] >= rep["start"]) & (tx_all["date"] <= rep["end"]) if not tx_all.empty else []
            csv_day = REPORT_DIR / f"report_{rep['label']}_day.csv"
            (tx_all.loc[mask] if not tx_all.empty else tx_all).to_csv(csv_day, index=False)
            atts.append(csv_day)

        html = render_html_summary(rep)
        print(f"[REPORT] {per} {rep['label']} | spent={rep['spent']:.2f} | atts={len(atts)}")
        if send:
            send_email(subject, html, atts)
            print(f"[MAIL] gesendet an: {EMAIL_TO}")

# ====================== CLI ======================
if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="Finanz-Reports als PDF per E-Mail senden")
    ap.add_argument("period", choices=["daily","monthly","quarterly","yearly","auto"], help="Zeitraum")
    ap.add_argument("--send", action="store_true", help="E-Mail wirklich senden (sonst nur erzeugen)")
    args = ap.parse_args()
    run(args.period, send=args.send)
