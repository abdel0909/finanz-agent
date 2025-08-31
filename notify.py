#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
notify.py – Tages/Monats/Quartals/Jahres-Report per E-Mail (Gmail SMTP)
- Repariert/legt Budget.xlsx bei Bedarf neu an
- Berechnet Kennzahlen & Zeitraum-Daten
- Erstellt Diagramme (matplotlib)
- Baut schlankes, farbiges PDF (ReportLab) mit schmaler Tabelle
- Sendet E-Mail (HTML + PDF + CSV-Anhänge)

Abhängigkeiten (requirements.txt):
  pandas
  openpyxl
  matplotlib
  reportlab
"""

import os, ssl, smtplib, re
from pathlib import Path
from datetime import datetime, date, timedelta
from zoneinfo import ZoneInfo
from zipfile import BadZipFile
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

import pandas as pd
import matplotlib.pyplot as plt

# ============ Pfade & Umgebung ============
BASE        = Path(__file__).parent.resolve()
WORKBOOK    = BASE / "Budget.xlsx"
REPORT_DIR  = BASE / "reports"
IMG_DIR     = REPORT_DIR / "imgs"
DEBUG_DIR   = REPORT_DIR / "debug"
TZ          = ZoneInfo(os.getenv("LOCAL_TZ", "Europe/Berlin"))

# Mail / Profile (Gmail App-Passwort nutzen!)
SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.gmail.com")
SMTP_PORT   = int(os.getenv("SMTP_PORT", "587"))
SMTP_USER   = os.getenv("SMTP_USER", "bouardjaa@gmail.com")
SMTP_PASS   = os.getenv("SMTP_PASS", "zwqdwuyxdzydtaqu")
EMAIL_FROM  = os.getenv("bouardjaa@gmail.com", SMTP_USER)
EMAIL_TO    = [a.strip() for a in os.getenv("bouardjaa@gmail.com", SMTP_USER).split(",") if a.strip()]

PROFILE_NAME    = os.getenv("PROFILE_NAME", "")       # z.B. "Max Mustermann"
PROFILE_ADDRESS = os.getenv("PROFILE_ADDRESS", "")    # z.B. "Musterstraße 1, 12345 Musterstadt"

# ============ Excel sichern / laden ============
def _write_empty_workbook() -> None:
    """Legt minimale Budget.xlsx mit nötigen Sheets an."""
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
    tx["category"] = tx.get("category", "").astype(str)
    return tx

# ============ Perioden / Kennzahlen ============
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
        label = f"{s.year}-Q{(s.month-1)//3 + 1}"
    elif period == "yearly":
        s = date(ref.year-1, 1, 1); e = date(ref.year-1, 12, 31); label = str(s.year)
    else:
        raise ValueError("period must be daily|monthly|quarterly|yearly")
    return s, e, label

def _sum_spent(df: pd.DataFrame) -> float:
    return float((-df.loc[df["amount_eur"] < 0, "amount_eur"]).sum())

def build_report(period: str, today: date | None = None) -> tuple[dict, pd.DataFrame]:
    tx = load_tx()
    if today is None:
        today = datetime.now(TZ).date()
    s, e, label = _period_range(period, today)
    mask = (tx["date"] >= s) & (tx["date"] <= e) if not tx.empty else []
    df = tx.loc[mask].copy() if not tx.empty else pd.DataFrame(columns=tx.columns)

    total_spent = _sum_spent(df) if not df.empty else 0.0
    income      = float(df.loc[df["amount_eur"] > 0, "amount_eur"].sum()) if not df.empty else 0.0
    net         = float(df["amount_eur"].sum()) if not df.empty else 0.0

    # Kategorien für Diagramm
    IMG_DIR.mkdir(parents=True, exist_ok=True)
    images: list[Path] = []
    if not df.empty:
        cats = df[df["amount_eur"] < 0].groupby("category", as_index=False)["amount_eur"].sum()
        if not cats.empty:
            cats["spent"] = -cats["amount_eur"]
            # Horizontales Balkendiagramm
            fig = plt.figure()
            top = cats.sort_values("spent", ascending=False).head(10).sort_values("spent")
            plt.barh(top["category"], top["spent"])
            plt.title(f"Top-Kategorien {period} {label}")
            plt.xlabel("EUR")
            p = IMG_DIR / f"{period}_{label}_categories.png"
            fig.tight_layout(); fig.savefig(p); plt.close(fig)
            images.append(p)

        # Zeitreihe kumulierte Ausgaben (nur für >= monatlich)
        if period in ("monthly","quarterly","yearly"):
            fig = plt.figure()
            day = df.groupby("date", as_index=False)["amount_eur"].sum()
            day["cum_spent"] = (-day["amount_eur"].clip(upper=0)).cumsum()
            plt.plot(day["date"], day["cum_spent"])
            plt.title(f"Kumulierte Ausgaben {period} {label}")
            plt.xlabel("Datum"); plt.ylabel("EUR")
            p = IMG_DIR / f"{period}_{label}_timeseries.png"
            fig.tight_layout(); fig.savefig(p); plt.close(fig)
            images.append(p)

    # einfache Bewertung
    savings_rate = (income + net) / income if income > 0 else 0.0
    rating = "✅ gut"
    if income == 0 and total_spent > 0: rating = "⚠️ nur Ausgaben"
    elif savings_rate < 0.05:          rating = "⚠️ sehr niedrig"
    elif savings_rate < 0.15:          rating = "🟡 mittel"

    rep = {
        "period": period, "label": label, "start": s, "end": e,
        "spent": round(total_spent, 2), "income": round(income, 2), "net": round(net, 2),
        "images": images, "rating": rating
    }
    return rep, df

# ============ PDF (ReportLab) ============
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (BaseDocTemplate, PageTemplate, Frame, Paragraph, Spacer,
                                Table, TableStyle, Image, Flowable)

class Bar(Flowable):
    """schmale farbige Leiste über dem Titel"""
    def __init__(self, color=colors.HexColor("#2064ff"), height=14):
        Flowable.__init__(self)
        self.color = color
        self.height = height
        self.width = 0
    def wrap(self, availWidth, availHeight):
        self.width = availWidth
        return availWidth, self.height
    def draw(self):
        c = self.canv
        c.setFillColor(self.color)
        c.setStrokeColor(self.color)
        c.rect(0, 0, self.width, self.height, stroke=0, fill=1)

def _styles():
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="TitleBig", parent=styles["Title"], fontSize=22, leading=26, alignment=1, spaceAfter=6))
    styles.add(ParagraphStyle(name="Sub", parent=styles["Normal"], fontSize=11, leading=14, alignment=1, textColor=colors.HexColor("#666")))
    styles.add(ParagraphStyle(name="H2", parent=styles["Heading2"], fontSize=14, leading=18, spaceBefore=8, spaceAfter=6))
    styles.add(ParagraphStyle(name="CardTitle", parent=styles["Normal"], fontSize=10, textColor=colors.HexColor("#4c4c4c")))
    styles.add(ParagraphStyle(name="CardValue", parent=styles["Heading2"], fontSize=16, leading=18))
    styles.add(ParagraphStyle(name="Small", parent=styles["Normal"], fontSize=8, textColor=colors.HexColor("#777")))
    styles.add(ParagraphStyle(name="Cell", parent=styles["Normal"], fontSize=8, leading=10))
    return styles

def _money(v: float) -> str:
    return f"{v:,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")

def _card(label, value, border=colors.black):
    s = _styles()
    box = Table(
        [[Paragraph(label, s["CardTitle"])],
         [Paragraph(_money(value), s["CardValue"])]],
        colWidths=[70*mm], rowHeights=[8*mm, 14*mm], hAlign="LEFT",
    )
    box.setStyle(TableStyle([
        ("BOX", (0,0), (-1,-1), 1, border),
        ("TOPPADDING", (0,0), (-1,-1), 2),
        ("BOTTOMPADDING", (0,0), (-1,-1), 2),
        ("ALIGN", (0,1), (-1,-1), "LEFT"),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
    ]))
    return box

def _table_from_df(df: pd.DataFrame) -> Table:
    s = _styles()
    show_cols = ["date","account_id","payee","amount_eur","currency","category","tags","note"]
    df2 = df.copy()
    for c in show_cols:
        if c not in df2.columns:
            df2[c] = ""

    df2["date"] = pd.to_datetime(df2["date"], errors="coerce").dt.date.astype(str)
    df2["amount_eur"] = df2["amount_eur"].map(lambda x: _money(x) if pd.notnull(x) else "")

    header = [Paragraph(col, s["Cell"]) for col in show_cols]
    rows = []
    for _, r in df2[show_cols].iterrows():
        row = [Paragraph("" if pd.isna(r[c]) else str(r[c]), s["Cell"]) for c in show_cols]
        rows.append(row)

    data = [header] + rows

    col_widths = [
        20*mm,  # date
        20*mm,  # account_id
        28*mm,  # payee
        20*mm,  # amount_eur
        12*mm,  # currency
        22*mm,  # category
        25*mm,  # tags
        28*mm,  # note
    ]

    tbl = Table(data, colWidths=col_widths, hAlign="LEFT", repeatRows=1)
    tbl.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#2064ff")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("ALIGN", (0,0), (-1,0), "LEFT"),
        ("FONTSIZE", (0,0), (-1,-1), 8),
        ("LEADING", (0,1), (-1,-1), 9.5),
        ("GRID", (0,0), (-1,-1), 0.5, colors.HexColor("#d0d0d0")),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.whitesmoke, colors.white]),
        ("BOTTOMPADDING", (0,0), (-1,-1), 3),
        ("TOPPADDING", (0,0), (-1,-1), 3),
    ]))
    return tbl

def render_pdf_statement(rep: dict, df_period: pd.DataFrame, pdf_path: Path, profile: dict | None = None):
    profile = profile or {}
    s = _styles()

    margin = 15*mm
    doc = BaseDocTemplate(str(pdf_path), pagesize=A4,
                          leftMargin=margin, rightMargin=margin,
                          topMargin=18*mm, bottomMargin=15*mm)
    frame = Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height, id="normal")
    doc.addPageTemplates([PageTemplate(id="main", frames=[frame])])

    elements = []
    # Kopf
    elements.append(Bar())
    elements.append(Spacer(1, 8*mm))
    if profile.get("name") or profile.get("address"):
        elements.append(Paragraph(profile.get("name",""), s["Small"]))
        elements.append(Paragraph(profile.get("address",""), s["Small"]))
        elements.append(Spacer(1, 2*mm))
    title = "Finanz-Report"
    subtitle = f"{rep['period'].capitalize()} – {rep['label']}<br/>{rep['start']} bis {rep['end']}"
    elements.append(Paragraph(title, s["TitleBig"]))
    elements.append(Paragraph(subtitle, s["Sub"]))
    elements.append(Spacer(1, 6*mm))

    # Kennzahlen
    cards = Table(
        [[_card("Ausgaben", -abs(rep["spent"]), border=colors.HexColor("#1f6bff")),
          _card("Einnahmen", rep["income"], border=colors.HexColor("#31a24c")),
          _card("Netto", rep["net"], border=colors.black)]],
        colWidths=[65*mm, 65*mm, 45*mm],
        hAlign="LEFT",
    )
    cards.setStyle(TableStyle([("VALIGN", (0,0), (-1,-1), "MIDDLE")]))
    elements.append(cards)
    elements.append(Spacer(1, 8*mm))

    # Tabelle
    elements.append(Paragraph("Buchungen", s["H2"]))
    if df_period is not None and not df_period.empty:
        elements.append(_table_from_df(df_period))
        elements.append(Spacer(1, 5*mm))
        elements.append(Paragraph(f"Erstellt am {datetime.now(TZ).strftime('%Y-%m-%d %H:%M')}", s["Small"]))
    else:
        elements.append(Paragraph("(Keine Buchungen im Zeitraum.)", s["Small"]))

    # Diagramm (erstes vorhandenes Bild)
    img_path = None
    if rep.get("images"):
        img_path = rep["images"][0]
    if img_path and Path(img_path).exists():
        elements.append(Spacer(1, 10*mm))
        elements.append(Paragraph("Diagramm", s["H2"]))
        img = Image(str(img_path))
        max_w = doc.width
        img.drawWidth  = max_w
        img.drawHeight = max_w * img.imageHeight / img.imageWidth
        elements.append(img)

    doc.build(elements)

# ============ HTML für E-Mail ============
def render_html(rep: dict) -> str:
    def fmt(v): return f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"""
    <h2>Finanz-Report ({rep['period']} – {rep['label']})</h2>
    <p><b>Zeitraum:</b> {rep['start']} bis {rep['end']}</p>
    <ul>
      <li><b>Ausgaben</b>: {fmt(rep['spent'])} €</li>
      <li><b>Einnahmen</b>: {fmt(rep['income'])} €</li>
      <li><b>Netto</b>: {fmt(rep['net'])} €</li>
      <li><b>Bewertung</b>: {rep['rating']}</li>
    </ul>
    <p>PDF im Anhang enthält Tabelle & Diagramm.</p>
    """

# ============ Mailversand ============
def _write_debug(html: str):
    DEBUG_DIR.mkdir(parents=True, exist_ok=True)
    (DEBUG_DIR / "last_mail.html").write_text(html, encoding="utf-8")
    (DEBUG_DIR / "last_mail.txt").write_text(re.sub(r"<[^>]+>", "", html), encoding="utf-8")

def send_email(subject: str, html: str, attachments: list[Path]) -> None:
    if not (SMTP_USER and SMTP_PASS and EMAIL_FROM and EMAIL_TO):
        raise RuntimeError("SMTP-Variablen fehlen (SMTP_USER, SMTP_PASS, EMAIL_FROM, EMAIL_TO).")

    if not html.strip():
        html = "<p>(Kein Inhalt)</p>"
    _write_debug(html)

    msg = MIMEMultipart("mixed")
    msg["Subject"] = subject
    msg["From"] = EMAIL_FROM
    msg["To"] = ",".join(EMAIL_TO)

    alt = MIMEMultipart("alternative")
    alt.attach(MIMEText(re.sub(r"<[^>]+>", "", html), "plain", "utf-8"))
    alt.attach(MIMEText(html, "html", "utf-8"))
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

# ============ Orchestrierung ============
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
            if today.month in (1,4,7,10): periods.append("quarterly")
            if today.month == 1: periods.append("yearly")
    else:
        periods = [period]

    for per in periods:
        rep, df = build_report(per, today=today)

        # Dateiname & CSV für Perioden-Daten
        label_safe = rep["label"]
        if per == "daily":
            csv_file = REPORT_DIR / f"report_{label_safe}_day.csv"
            (df if not df.empty else pd.DataFrame()).to_csv(csv_file, index=False)
        elif per == "monthly":
            csv_file = REPORT_DIR / f"report_{label_safe}.csv"
            if not csv_file.exists():
                (df if not df.empty else pd.DataFrame()).to_csv(csv_file, index=False)
        else:
            csv_file = None

        # PDF erzeugen
        pdf_name = {
            "daily":     f"statement_{label_safe}.pdf",
            "monthly":   f"statement_{label_safe}.pdf",
            "quarterly": f"statement_{label_safe}.pdf",
            "yearly":    f"statement_{label_safe}.pdf",
        }[per]
        pdf_path = REPORT_DIR / pdf_name
        render_pdf_statement(rep, df, pdf_path, profile={"name": PROFILE_NAME, "address": PROFILE_ADDRESS})

        # E-Mail
        subject = f"Finanzen – {per} – {rep['label']} (Ausgaben {rep['spent']:.2f} €)"
        html = render_html(rep)
        atts = [pdf_path] + ([csv_file] if csv_file else [])
        if send:
            send_email(subject, html, atts)
            print(f"[MAIL] gesendet an: {EMAIL_TO}")
        print(f"[REPORT] {per} {rep['label']} | spent={rep['spent']:.2f} | atts={len(atts)}")

# ============ CLI ============
if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="Finanz-Reports per E-Mail senden")
    ap.add_argument("period", choices=["daily","monthly","quarterly","yearly","auto"], help="Zeitraum")
    ap.add_argument("--send", action="store_true", help="E-Mail wirklich senden (sonst nur erzeugen)")
    args = ap.parse_args()
    run(args.period, send=args.send)
    
