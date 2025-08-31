#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
notify.py – PDF/HTML-Report per Mail (Gmail SMTP)
Erzeugt/repariert Budget.xlsx, generiert Diagramme, baut schönes PDF
und verschickt alles per Mail.
"""

from __future__ import annotations
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

# --- ReportLab (PDF)
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
)
from reportlab.platypus.flowables import Flowable

# ======================= Konfiguration =======================
BASE       = Path(__file__).parent.resolve()
WORKBOOK   = BASE / "Budget.xlsx"
REPORT_DIR = BASE / "reports"
IMG_DIR    = REPORT_DIR / "imgs"
TZ         = ZoneInfo(os.getenv("LOCAL_TZ", "Europe/Berlin"))

# Mail/Gmail – per ENV oder GitHub Secrets setzen
SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.gmail.com")
SMTP_PORT   = int(os.getenv("SMTP_PORT", "587"))
SMTP_USER   = os.getenv("SMTP_USER", "bouardjaa@gmail.com")
SMTP_PASS   = os.getenv("SMTP_PASS", "zwqdwuyxdzydtaqu")
EMAIL_FROM  = os.getenv("bouardjaa@gmail.com", SMTP_USER or "noreply@example.com")
EMAIL_TO    = [a.strip() for a in os.getenv("bouardjaa@gmail.com", SMTP_USER).split(",") if a.strip()]

# Briefkopf (optional)
PROFILE_NAME    = os.getenv("PROFILE_NAME", "")     # z.B. "Max Mustermann"
PROFILE_ADDRESS = os.getenv("PROFILE_ADDRESS", "")  # z.B. "Musterstr. 1, 12345 Musterstadt"

# ======================= Excel-Helfer ========================
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

# ======================= Perioden/Report =====================
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
        last_prev = first_this_q - timedelta(days=1)
        qm = ((last_prev.month - 1)//3)*3 + 1
        s = date(last_prev.year, qm, 1)
        e = date(last_prev.year, qm+3, 1) - timedelta(days=1)
        label = f"{s.strftime('%Y')}-Q{((s.month-1)//3)+1}"
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

    # Kategorien & Diagramme
    cats = pd.DataFrame(columns=["category","spent"])
    images: list[Path] = []
    IMG_DIR.mkdir(parents=True, exist_ok=True)

    if not df.empty:
        cat = df[df["amount_eur"] < 0].groupby("category", as_index=False)["amount_eur"].sum()
        cat["spent"] = -cat["amount_eur"]
        cats = cat[["category","spent"]].sort_values("spent", ascending=False)

        # Bar-Chart Top-Kategorien (keine Farb-Styles vorgeben)
        if not cats.empty:
            fig = plt.figure()
            top = cats.head(10).sort_values("spent")
            plt.barh(top["category"], top["spent"])
            plt.title(f"Top-Kategorien {period} {label}")
            plt.xlabel("EUR")
            p = IMG_DIR / f"{period}_{label}_categories.png"
            fig.tight_layout(); fig.savefig(p); plt.close(fig)
            images.append(p)

    # Bewertung (simple savings rate)
    savings_rate = (income + net) / income if income > 0 else 0.0
    rating = "✅ gut"
    if income == 0 and total_spent > 0: rating = "⚠️ nur Ausgaben"
    elif savings_rate < 0.05:          rating = "⚠️ sehr niedrig"
    elif savings_rate < 0.15:          rating = "🟡 mittel"

    return {
        "period": period, "label": label, "start": s, "end": e,
        "spent": round(total_spent, 2), "income": round(income, 2), "net": round(net, 2),
        "top_categories": cats.to_dict("records"), "images": images, "rating": rating,
        "df": df
    }

# ======================= HTML-Body ===========================
def render_html(rep: dict) -> str:
    cats = "".join([f"<li>{c['category']}: {c['spent']:.2f} €</li>" for c in rep["top_categories"][:10]]) or "<li>(keine)</li>"
    return f"""
    <h2>Finanz-Report ({rep['period']} – {rep['label']})</h2>
    <p><b>Zeitraum:</b> {rep['start']} bis {rep['end']}</p>
    <ul>
      <li><b>Ausgaben</b>: {rep['spent']:.2f} €</li>
      <li><b>Einnahmen</b>: {rep['income']:.2f} €</li>
      <li><b>Netto</b>: {rep['net']:.2f} €</li>
      <li><b>Bewertung</b>: {rep['rating']}</li>
    </ul>
    <p><b>Top-Kategorien (Ausgaben)</b></p>
    <ul>{cats}</ul>
    """

# ======================= PDF (kompakt & sauber) ==============
def _chunk(lst, n):
    for i in range(0, len(lst), n):
        yield lst[i:i+n]

def _auto_col_widths(df, headers, avail_width):
    mins = {
        "date": 20*mm, "account_id": 28*mm, "payee": 35*mm, "amount_eur": 22*mm,
        "currency": 16*mm, "category": 28*mm, "tags": 28*mm, "note": 40*mm
    }
    maxs = {
        "date": 28*mm, "account_id": 38*mm, "payee": 65*mm, "amount_eur": 30*mm,
        "currency": 20*mm, "category": 45*mm, "tags": 50*mm, "note": 90*mm
    }
    char_mm = 2.3
    lengths = []
    for h in headers:
        col = df[h].astype(str).fillna("")
        longest = max([len(h)] + [len(s) for s in col.head(300)])
        lengths.append(longest)
    widths = [max(mins.get(h, 18*mm), min(maxs.get(h, 70*mm), l*char_mm)) for h,l in zip(headers, lengths)]
    scale = avail_width / sum(widths)
    widths = [w*scale for w in widths]
    fixed=[]
    for w,h in zip(widths, headers):
        w = max(w, mins.get(h, 18*mm))
        w = min(w, maxs.get(h, 90*mm))
        fixed.append(w)
    rest = avail_width - sum(fixed)
    if rest > 0 and "note" in headers:
        i = headers.index("note"); fixed[i] += rest
    return fixed

class Bar(Flowable):
    def __init__(self, color=colors.HexColor("#2064ff"), height=14):
        Flowable.__init__(self)
        self.color = color
        self.height = height
        self.width = 0  # wird in wrap gesetzt

    def wrap(self, availWidth, availHeight):
        self.width = availWidth
        return availWidth, self.height

    def draw(self):
        c = self.canv
        c.setFillColor(self.color)
        c.setStrokeColor(self.color)
        c.rect(0, 0, self.width, self.height, stroke=0, fill=1)

def render_pdf_statement(rep: dict, out_pdf: Path, profile: dict | None = None):
    df = rep["df"].copy()
    profile = profile or {}
    name    = profile.get("name", "")
    address = profile.get("address", "")

    page_w, page_h = A4
    margin = 18*mm
    doc = SimpleDocTemplate(
        str(out_pdf),
        pagesize=A4,
        leftMargin=margin, rightMargin=margin, topMargin=16*mm, bottomMargin=16*mm
    )

    # Styles
    h1 = ParagraphStyle("h1", fontName="Helvetica-Bold", fontSize=18, leading=22, alignment=1, spaceAfter=6)
    h2 = ParagraphStyle("h2", fontName="Helvetica", fontSize=12, leading=16, alignment=1, textColor=colors.grey)
    small = ParagraphStyle("small", fontName="Helvetica", fontSize=9, leading=12, textColor=colors.grey)
    body = ParagraphStyle("body", fontName="Helvetica", fontSize=10, leading=12)
    wrap = ParagraphStyle("wrap", fontName="Helvetica", fontSize=9, leading=11)

    elements = []

    # Kopf
    if name or address:
        txt = f"<b>{name}</b><br/>{address}" if address else f"<b>{name}</b>"
        elements.append(Paragraph(txt, small)); elements.append(Spacer(1, 6))

    elements.append(Bar()); elements.append(Spacer(1, 10))
    title = "Finanz-Report"
    subtitle = f"{rep['period'].title()} – {rep['label']}<br/>{rep['start']} bis {rep['end']}"
    elements.append(Paragraph(title, h1))
    elements.append(Paragraph(subtitle, h2))
    elements.append(Spacer(1, 10))

    # KPI-Karten
    def _kpi(label, value, color):
        lab = Paragraph(f"<font size=9 color='{color.rgb()}'>" + label + "</font>", body)
        val = Paragraph(f"<para alignment='left'><b>{value:,.2f} €</b></para>",
                        ParagraphStyle("v", fontName="Helvetica-Bold", fontSize=14, leading=16))
        box = Table([[lab],[val]], colWidths=[(page_w-2*margin)/3 - 6])
        box.setStyle(TableStyle([
            ("BOX", (0,0), (-1,-1), 1, color),
            ("LEFTPADDING", (0,0), (-1,-1), 6), ("RIGHTPADDING", (0,0), (-1,-1), 6),
            ("TOPPADDING", (0,0), (-1,-1), 4), ("BOTTOMPADDING", (0,0), (-1,-1), 6),
        ]))
        return box
    kpis = Table([[
        _kpi("Ausgaben", -float(rep["spent"]), colors.HexColor("#1f6bff")),
        _kpi("Einnahmen", float(rep["income"]), colors.HexColor("#2aa164")),
        _kpi("Netto", float(rep["net"]), colors.black),
    ]], colWidths=[(page_w-2*margin)/3 - 6]*3, hAlign="LEFT",
       style=TableStyle([("VALIGN",(0,0),(-1,-1),"MIDDLE")]))
    elements.append(kpis); elements.append(Spacer(1, 12))

    elements.append(Paragraph("<b>Buchungen</b>", ParagraphStyle("b", fontName="Helvetica-Bold", fontSize=12, leading=14)))
    elements.append(Spacer(1, 4))

    if df.empty:
        elements.append(Paragraph("(keine Buchungen im Zeitraum)", body))
    else:
        keep = ["date","account_id","payee","amount_eur","currency","category","tags","note"]
        df2 = df[keep].copy()
        df2["date"] = pd.to_datetime(df2["date"], errors="coerce").dt.date.astype(str)
        df2["amount_eur"] = df2["amount_eur"].map(
            lambda x: f"{x:,.2f} €".replace(",", "X").replace(".", ",").replace("X",".")
        )
        headers = keep
        labels  = ["date","account_id","payee","amount_eur","currency","category","tags","note"]

        avail_width = page_w - 2*margin
        col_widths = _auto_col_widths(df2, headers, avail_width)

        def _row(rec):
            row=[]
            for h in headers:
                txt = str(rec[h]) if pd.notna(rec[h]) else ""
                if h in ("date","account_id","amount_eur","currency","category"):
                    row.append(Paragraph(txt, body))
                else:
                    row.append(Paragraph(txt, wrap))
            return row

        data = [labels] + [_row(r) for _, r in df2.iterrows()]
        max_rows = 28  # pro Seite (Header zählt nicht)

        for part in _chunk(data[1:], max_rows):
            tbl = Table([labels] + part, colWidths=col_widths, repeatRows=1)
            tbl.setStyle(TableStyle([
                ("GRID", (0,0), (-1,-1), 0.25, colors.grey),
                ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#1f6bff")),
                ("TEXTCOLOR", (0,0), (-1,0), colors.white),
                ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
                ("FONTSIZE", (0,0), (-1,0), 9),
                ("LEADING", (0,0), (-1,0), 11),
                ("VALIGN", (0,0), (-1,-1), "TOP"),
                ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.whitesmoke, colors.Color(0,0,0,0)]),
                ("LEFTPADDING", (0,0), (-1,-1), 4), ("RIGHTPADDING", (0,0), (-1,-1), 4),
                ("TOPPADDING", (0,0), (-1,-1), 3), ("BOTTOMPADDING", (0,0), (-1,-1), 4),
            ]))
            elements.append(tbl)
            elements.append(Spacer(1, 6))
            elements.append(Paragraph(f"Erstellt am {datetime.now().strftime('%Y-%m-%d %H:%M')}", small))
            elements.append(PageBreak())
        if isinstance(elements[-1], PageBreak):
            elements.pop()

    def _footer(canvas, doc):
        canvas.saveState()
        canvas.setFont("Helvetica", 9)
        canvas.setFillColor(colors.grey)
        canvas.drawRightString(page_w - margin, 10*mm, f"Seite {doc.page}")
        canvas.restoreState()

    out_pdf.parent.mkdir(parents=True, exist_ok=True)
    doc.build(elements, onLaterPages=_footer, onFirstPage=_footer)

# ======================= Mailversand =========================
def _write_debug(html: str):
    d = REPORT_DIR / "debug"; d.mkdir(parents=True, exist_ok=True)
    (d/"last_mail.html").write_text(html or "", encoding="utf-8")
    (d/"last_mail.txt").write_text(re.sub(r"<[^>]+>", "", html or ""), encoding="utf-8")

def send_email(subject: str, html: str, attachments: list[Path]) -> None:
    if not (SMTP_USER and SMTP_PASS and EMAIL_FROM and EMAIL_TO):
        raise RuntimeError("SMTP/Gmail Variablen fehlen (SMTP_USER/SMTP_PASS/EMAIL_FROM/EMAIL_TO).")
    if not html.strip():
        html = "<p>(Kein HTML-Inhalt erzeugt)</p>"
    _write_debug(html)

    msg = MIMEMultipart("mixed")
    msg["Subject"] = subject
    msg["From"] = EMAIL_FROM
    msg["To"] = ",".join(EMAIL_TO)

    alt = MIMEMultipart("alternative")
    alt.attach(MIMEText(re.sub(r"<[^>]+>", "", html).strip() or "(Kein Inhalt)", "plain", "utf-8"))
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

# ======================= Orchestrierung =====================
def _daily_at_20_local() -> bool:
    return datetime.now(TZ).hour == 20

def run(period: str, send: bool = True) -> None:
    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    today = datetime.now(TZ).date()

    periods: list[str] = []
    if period == "auto":
        if _daily_at_20_local(): periods.append("daily")
        if today.day == 1:
            periods += ["monthly"]
            if today.month in (1,4,7,10): periods.append("quarterly")
            if today.month == 1: periods.append("yearly")
    else:
        periods = [period]

    for per in periods:
        rep = build_report(per, today=today)
        html = render_html(rep)

        # CSV-Anhang fürs Detail (nur daily/monat bereits erzeugt)
        atts: list[Path] = []
        if not rep["df"].empty:
            if per == "daily":
                dayfile = REPORT_DIR / f"report_{rep['label']}_day.csv"
                rep["df"].to_csv(dayfile, index=False); atts.append(dayfile)
            elif per == "monthly":
                mon = REPORT_DIR / f"report_{rep['label']}.csv"
                if mon.exists(): atts.append(mon)

        # PDF erzeugen
        pdf_path = REPORT_DIR / f"statement_{rep['label']}.pdf"
        render_pdf_statement(
            rep,
            out_pdf=pdf_path,
            profile={"name": PROFILE_NAME, "address": PROFILE_ADDRESS}
        )
        atts.insert(0, pdf_path)  # PDF zuerst

        subject = f"Finanzen – {per} – {rep['label']} (Ausgaben {rep['spent']:.2f} €)"
        print(f"[REPORT] {per} {rep['label']} | spent={rep['spent']:.2f} | atts={len(atts)} | pdf={pdf_path.name}")

        if send:
            send_email(subject, html, atts)
            print(f"[MAIL] gesendet an: {EMAIL_TO}")

# ======================= CLI ================================
if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="Finanz-Reports per E-Mail senden (inkl. PDF)")
    ap.add_argument("period", choices=["daily","monthly","quarterly","yearly","auto"], help="Zeitraum")
    ap.add_argument("--send", action="store_true", help="E-Mail wirklich senden (sonst nur Dateien erzeugen)")
    args = ap.parse_args()
    run(args.period, send=args.send)
