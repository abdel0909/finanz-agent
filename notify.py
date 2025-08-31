#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
notify.py – PDF- und Mail-Reports für Finanz-Agent
Erzeugt Diagramme, baut ein sauberes PDF (A4, Farblayout) und verschickt E-Mails via Gmail SMTP.

ENV (Beispiele):
  LOCAL_TZ="Europe/Berlin"
  SMTP_SERVER="smtp.gmail.com"
  SMTP_PORT="587"
  SMTP_USER="dein.gmail@gmail.com"
  SMTP_PASS="DEIN_APP_PASSWORT"
  EMAIL_FROM="dein.gmail@gmail.com"
  EMAIL_TO="ziel1@gmail.com,ziel2@gmail.com"

  PROFILE_NAME="Max Mustermann"
  PROFILE_ADDRESS="Musterstraße 1, 12345 Musterstadt"
"""

from __future__ import annotations
import os, ssl, smtplib, re
from pathlib import Path
from datetime import datetime, date, timedelta
from zoneinfo import ZoneInfo
from zipfile import BadZipFile
from typing import List, Dict, Tuple

import pandas as pd
import matplotlib.pyplot as plt

# === ReportLab für PDF ===
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage, KeepTogether
)

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

PROFILE_NAME    = os.getenv("PROFILE_NAME", "")
PROFILE_ADDRESS = os.getenv("PROFILE_ADDRESS", "")

# ------------------- Helpers -------------------
def fmt_eur(x: float) -> str:
    """Deutsch formatiert, z. B. -1.5 -> '-1,50 €'."""
    try:
        s = f"{x:,.2f}"
    except Exception:
        s = "0.00"
    s = s.replace(",", "_").replace(".", ",").replace("_", ".")
    return f"{s} €"

def write_debug(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")

# ------------------- Excel / Daten -------------------
def _write_empty_workbook() -> None:
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
    if not WORKBOOK.exists():
        _write_empty_workbook(); return
    try:
        pd.read_excel(WORKBOOK, sheet_name="Transactions", engine="openpyxl")
    except BadZipFile:
        print("[WARN] Budget.xlsx beschädigt → neu anlegen …")
        WORKBOOK.unlink(missing_ok=True)
        _write_empty_workbook()

def load_all_tx() -> pd.DataFrame:
    _ensure_workbook_ok()
    df = pd.read_excel(WORKBOOK, sheet_name="Transactions", engine="openpyxl")
    if df.empty:
        return df
    df["date"] = pd.to_datetime(df["date"], errors="coerce").dt.date
    df["amount_eur"] = pd.to_numeric(df["amount_eur"], errors="coerce").fillna(0.0)
    for col in ("account_id","payee","currency","category","tags","note"):
        if col in df.columns:
            df[col] = df[col].astype(str)
    return df

# ------------------- Perioden/Reporting -------------------
def period_range(period: str, ref: date) -> Tuple[date, date, str]:
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
        s = date(ref.year-1, 1, 1); e = date(ref.year-1, 12, 31); label = f"{s.year}"
    else:
        raise ValueError("period must be daily|monthly|quarterly|yearly")
    return s, e, label

def build_report(period: str, today: date | None = None) -> Dict:
    if today is None:
        today = datetime.now(TZ).date()
    s, e, label = period_range(period, today)

    tx_all = load_all_tx()
    if tx_all.empty:
        df = pd.DataFrame(columns=[
            "date","account_id","payee","amount_eur","currency",
            "category","tags","note","external_id","source"
        ])
    else:
        mask = (tx_all["date"] >= s) & (tx_all["date"] <= e)
        df = tx_all.loc[mask].copy()

    spent  = float((-df.loc[df["amount_eur"] < 0, "amount_eur"]).sum()) if not df.empty else 0.0
    income = float( (df.loc[df["amount_eur"] > 0, "amount_eur"]).sum()) if not df.empty else 0.0
    net    = float(df["amount_eur"].sum()) if not df.empty else 0.0

    # Top-Kategorien + Diagramm
    images: List[Path] = []
    (REPORT_DIR/"imgs").mkdir(parents=True, exist_ok=True)
    cats = pd.DataFrame(columns=["category","spent"])
    if not df.empty:
        cat = df[df["amount_eur"] < 0].groupby("category", as_index=False)["amount_eur"].sum()
        cat["spent"] = -cat["amount_eur"]
        cats = cat[["category","spent"]].sort_values("spent", ascending=False)

        if not cats.empty:
            fig = plt.figure(figsize=(7.5, 2.2))  # schmaler Balken
            top = cats.head(8).sort_values("spent")
            plt.barh(top["category"], top["spent"])
            plt.title(f"Top-Kategorien {period} {label}")
            plt.xlabel("EUR")
            plt.tight_layout()
            p = REPORT_DIR/"imgs"/f"{period}_{label}_categories.png"
            fig.savefig(p, dpi=160)
            plt.close(fig)
            images.append(p)

    # Bewertung
    savings_rate = (income + net) / income if income > 0 else 0.0
    rating = "✅ gut"
    if income == 0 and spent > 0: rating = "⚠️ nur Ausgaben"
    elif savings_rate < 0.05:     rating = "⚠️ sehr niedrig"
    elif savings_rate < 0.15:     rating = "🟡 mittel"

    return {
        "period": period, "label": label, "start": s, "end": e,
        "spent": round(spent, 2), "income": round(income, 2), "net": round(net, 2),
        "top_categories": cats.to_dict("records"),
        "images": images,
        "rating": rating,
        "df": df
    }

# ------------------- PDF -------------------
def render_pdf_statement(rep: Dict, out_pdf: Path, profile: Dict[str,str]) -> Path:
    """Schönes A4-PDF mit Kopf, KPIs, Tabelle, kleinem Diagramm."""
    REPORT_DIR.mkdir(parents=True, exist_ok=True)

    doc = SimpleDocTemplate(
        str(out_pdf),
        pagesize=A4,
        leftMargin=18*mm, rightMargin=18*mm, topMargin=18*mm, bottomMargin=18*mm
    )
    styles = getSampleStyleSheet()
    H = ParagraphStyle("H", parent=styles["Heading1"], alignment=1, fontSize=22, leading=26)
    Sub = ParagraphStyle("Sub", parent=styles["Normal"], alignment=1, textColor=colors.HexColor("#1a3e8a"))
    Label = ParagraphStyle("Label", parent=styles["Normal"], textColor=colors.HexColor("#666"))
    KPI = ParagraphStyle("KPI", parent=styles["Heading2"], fontSize=16, leading=18)
    Small = ParagraphStyle("Small", parent=styles["Normal"], fontSize=9, textColor=colors.HexColor("#666"))

    elements: List = []

    # Top-Leiste
    bar = Table([[" "]], colWidths=[doc.width], rowHeights=[8])
    bar.setStyle(TableStyle([("BACKGROUND",(0,0),(0,0), colors.HexColor("#1967ff"))]))
    elements += [bar, Spacer(1, 6)]

    # Titel
    elements += [
        Paragraph("Finanz-Report", H),
        Paragraph(f"{rep['period'].capitalize()} – {rep['label']}", Sub),
        Paragraph(f"{rep['start']} bis {rep['end']}", Label),
        Spacer(1, 8)
    ]

    # Profil (optional)
    if profile.get("name") or profile.get("address"):
        prof = Paragraph(
            f"{profile.get('name','')}"
            + (f"<br/>{profile.get('address','')}" if profile.get("address") else ""),
            styles["Normal"]
        )
        elements += [prof, Spacer(1, 6)]

    # KPI Boxen (3 Spalten)
    kpi_data = [
        [Paragraph("Ausgaben", Label), Paragraph("Einnahmen", Label), Paragraph("Netto", Label)],
        [Paragraph(fmt_eur(-rep["spent"]*-1) if rep["spent"]<0 else fmt_eur(rep["spent"]), KPI),
         Paragraph(fmt_eur(rep["income"]), KPI),
         Paragraph(fmt_eur(rep["net"]), KPI)]
    ]
    kpi = Table(kpi_data, colWidths=[doc.width/3-4, doc.width/3-4, doc.width/3-4])
    kpi.setStyle(TableStyle([
        ("BOX", (0,0), (-1,-1), 1, colors.black),
        ("INNERGRID", (0,0), (-1,-1), 0.6, colors.black),
        ("LINEABOVE",(0,0),(0,0),2, colors.HexColor("#1e66f5")),
        ("LINEABOVE",(1,0),(1,0),2, colors.HexColor("#12a150")),
        ("LINEABOVE",(2,0),(2,0),2, colors.black),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("TOPPADDING",(0,1),(-1,1),6),
        ("BOTTOMPADDING",(0,1),(-1,1),6),
    ]))
    elements += [kpi, Spacer(1, 10)]

    # Buchungstabelle (schmalere, feste Spaltenbreiten; Wrapping verhindert Überschlag)
    df = rep["df"].copy()
    if not df.empty:
        show_cols = ["date","account_id","payee","amount_eur","currency","category","tags","note"]
        df = df[show_cols].copy()
        df["amount_eur"] = df["amount_eur"].map(fmt_eur)

        data = [ [Paragraph(c, styles["BodyText"]) for c in show_cols] ]
        for _, row in df.iterrows():
            data.append([Paragraph(str(row[c]), styles["BodyText"]) for c in show_cols])

        # Spaltenbreiten in mm (werden auf Seitenbreite skaliert)
        target_mm = [20, 22, 35, 22, 16, 28, 28, 40]
        scale = (doc.width) / sum(w*mm for w in target_mm)
        col_widths = [w*mm*scale for w in target_mm]

        tbl = Table(data, colWidths=col_widths, repeatRows=1)
        tbl.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,0), colors.HexColor("#1e66f5")),
            ("TEXTCOLOR",(0,0),(-1,0), colors.white),
            ("ALIGN",(3,1),(3,-1),"RIGHT"),
            ("GRID",(0,0),(-1,-1),0.3, colors.black),
            ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
            ("FONTSIZE",(0,0),(-1,-1),9),
            ("ROWBACKGROUNDS",(0,1),(-1,-1), [colors.whitesmoke, colors.Color(1,1,1)])
        ]))
        elements += [Paragraph("Buchungen", styles["Heading2"]), tbl, Spacer(1, 6),
                     Paragraph(f"Erstellt am {datetime.now(TZ).strftime('%Y-%m-%d %H:%M')}", Small),
                     Spacer(1, 10)]
    else:
        elements += [Paragraph("Buchungen", styles["Heading2"]),
                     Paragraph("(Keine Buchungen im Zeitraum.)", styles["Normal"]),
                     Spacer(1, 10)]

    # Diagramm (klein, direkt unter Buchungen)
    if rep["images"]:
        elements += [Paragraph("Diagramm", styles["Heading2"])]
        # nur das erste Diagramm nutzen
        img_path = rep["images"][0]
        img = RLImage(str(img_path))
        img.drawWidth = doc.width * 0.85
        img.drawHeight = img.drawWidth * 0.45  # Verhältnis
        elements += [KeepTogether([img])]

    doc.build(elements)
    print(f"[PDF] geschrieben: {out_pdf.name}")
    return out_pdf

# ------------------- Email -------------------
def send_email(subject: str, html: str, attachments: List[Path]) -> None:
    if not (SMTP_USER and SMTP_PASS and EMAIL_FROM and EMAIL_TO):
        raise RuntimeError("SMTP-Variablen fehlen (SMTP_USER/SMTP_PASS/EMAIL_FROM/EMAIL_TO).")

    # Plaintext-Fallback
    plain = re.sub(r"<[^>]+>", "", html).replace("&nbsp;", " ").strip() or "(Kein Inhalt)"

    # Debug sichern
    dbg = REPORT_DIR/"debug"
    dbg.mkdir(parents=True, exist_ok=True)
    write_debug(dbg/"last_mail.html", html)
    write_debug(dbg/"last_mail.txt", plain)

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
        try:
            with open(p, "rb") as f:
                part = MIMEApplication(f.read(), Name=p.name)
            part["Content-Disposition"] = f'attachment; filename="{p.name}"'
            msg.attach(part)
        except Exception as ex:
            print(f"[WARN] Anhang übersprungen: {p} ({ex})")

    ctx = ssl.create_default_context()
    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as s:
        s.starttls(context=ctx)
        s.login(SMTP_USER, SMTP_PASS)
        s.sendmail(EMAIL_FROM, EMAIL_TO, msg.as_string())
    print(f"[MAIL] gesendet an: {EMAIL_TO}")

# ------------------- Orchestrierung -------------------
def daily_at_20_local() -> bool:
    return datetime.now(TZ).hour == 20

def run(period: str, send: bool = True) -> None:
    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    today = datetime.now(TZ).date()

    periods: List[str] = []
    if period == "auto":
        if daily_at_20_local(): periods.append("daily")
        if today.day == 1:
            periods.append("monthly")
            if today.month in (1,4,7,10): periods.append("quarterly")
            if today.month == 1: periods.append("yearly")
    else:
        periods = [period]

    for per in periods:
        rep = build_report(per, today=today)

        # CSV (Tagesliste) anlegen
        atts: List[Path] = []
        df = rep["df"]
        if per == "daily":
            dayfile = REPORT_DIR / f"report_{rep['label']}_day.csv"
            (df if not df.empty else df).to_csv(dayfile, index=False)
            atts.append(dayfile)

        # PDF erzeugen
        pdf_path = REPORT_DIR / f"statement_{rep['label']}.pdf"
        profile = {"name": PROFILE_NAME, "address": PROFILE_ADDRESS}
        render_pdf_statement(rep, pdf_path, profile)
        atts.append(pdf_path)

        # HTML für E-Mail (kurz)
        k_top = "".join([f"<li>{c['category']}: {fmt_eur(c['spent'])}</li>" for c in rep["top_categories"][:6]]) or "<li>(keine)</li>"
        html = f"""
        <h2>Finanz-Report ({rep['period']} – {rep['label']})</h2>
        <p><b>Zeitraum:</b> {rep['start']} bis {rep['end']}</p>
        <ul>
          <li><b>Ausgaben</b>: {fmt_eur(rep['spent'])}</li>
          <li><b>Einnahmen</b>: {fmt_eur(rep['income'])}</li>
          <li><b>Netto</b>: {fmt_eur(rep['net'])}</li>
          <li><b>Bewertung</b>: {rep['rating']}</li>
        </ul>
        <p><b>Top-Kategorien</b></p>
        <ul>{k_top}</ul>
        <p>PDF & CSV im Anhang.</p>
        """

        subject = f"Finanzen – {per} – {rep['label']} (Ausgaben {fmt_eur(rep['spent'])})"
        print(f"[REPORT] {per} {rep['label']} | spent={rep['spent']:.2f} | atts={len(atts)}")

        if send:
            send_email(subject, html, atts)

# ------------------- CLI -------------------
if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="Finanz-Reports per E-Mail (PDF/CSV) senden")
    ap.add_argument("period", choices=["daily","monthly","quarterly","yearly","auto"], help="Zeitraum / Modus")
    ap.add_argument("--send", action="store_true", help="E-Mail wirklich senden (sonst nur erzeugen)")
    args = ap.parse_args()
    run(args.period, send=args.send)
