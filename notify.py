#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
notify.py – Finanz-Reports per E-Mail (Gmail SMTP) + farbiges PDF
- Repariert/legt Budget.xlsx an
- Baut Diagramm-PNGs
- Sendet HTML-Mail
- Erstellt farbiges PDF (Urkunden-Layout) mit Tabelle über mehrere Seiten
"""

from __future__ import annotations
import os, ssl, smtplib, re, csv
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from pathlib import Path
from datetime import datetime, date, timedelta
from zoneinfo import ZoneInfo
from zipfile import BadZipFile

import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import matplotlib.image as mpimg

# ------------------- Pfade & Umgebung -------------------
BASE = Path(__file__).parent.resolve()
WORKBOOK = BASE / "Budget.xlsx"
REPORT_DIR = BASE / "reports"
IMG_DIR = REPORT_DIR / "imgs"
DEBUG_DIR = REPORT_DIR / "debug"
TZ = ZoneInfo(os.getenv("LOCAL_TZ", "Europe/Berlin"))

# SMTP / Mail
SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.gmail.com")
SMTP_PORT   = int(os.getenv("SMTP_PORT", "587"))
SMTP_USER   = os.getenv("SMTP_USER", "bouardjaa@gmail.com")
SMTP_PASS   = os.getenv("SMTP_PASS", "zwqdwuyxdzydtaqu")
EMAIL_FROM  = os.getenv("bouardjaa@gmail.com", SMTP_USER)
EMAIL_TO    = [a.strip() for a in os.getenv("bouardjaa@gmail.com", SMTP_USER).split(",") if a.strip()]

# Report-Kopf
REPORT_NAME  = os.getenv("REPORT_NAME", "")
REPORT_ADDR1 = os.getenv("REPORT_ADDR1", "")
REPORT_ADDR2 = os.getenv("REPORT_ADDR2", "")
REPORT_ADDR3 = os.getenv("REPORT_ADDR3", "")
PROFILE_ADDR_LINES = [x for x in (REPORT_ADDR1, REPORT_ADDR2, REPORT_ADDR3) if x]

# Farbschema (anpassbar via ENV – Hex ohne # möglich)
COLOR_PRIMARY   = "#" + os.getenv("PDF_COLOR_PRIMARY",   "0F62FE").lstrip("#")  # blau
COLOR_ACCENT    = "#" + os.getenv("PDF_COLOR_ACCENT",    "42BE65").lstrip("#")  # grün
COLOR_TEXT      = "#" + os.getenv("PDF_COLOR_TEXT",      "222222").lstrip("#")
COLOR_MUTED     = "#" + os.getenv("PDF_COLOR_MUTED",     "666666").lstrip("#")
COLOR_TABLE_ODD = "#" + os.getenv("PDF_COLOR_TABLE_ODD", "F7F8FA").lstrip("#")
COLOR_TABLE_EVN = "#" + os.getenv("PDF_COLOR_TABLE_EVN", "FFFFFF").lstrip("#")
COLOR_BORDER    = "#" + os.getenv("PDF_COLOR_BORDER",    "DDDDDD").lstrip("#")

# ------------------- Helpers -------------------
def _fmt_eur(v) -> str:
    try:
        x = float(v)
    except Exception:
        return str(v)
    s = f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return s + " €"

def _chunks(df: pd.DataFrame, size: int):
    for i in range(0, len(df), size):
        yield df.iloc[i:i+size]

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
    print("[INFO] Neue Budget.xlsx angelegt.")

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
    tx["amount_eur"] = pd.to_numeric(tx["amount_eur"], errors="coerce").fillna(0.0).round(2)
    for c in ("account_id","payee","currency","category","tags","note","external_id","source"):
        if c in tx.columns:
            tx[c] = tx[c].astype(str)
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
        label = f"{s.year}-Q{((s.month-1)//3)+1}"
    elif period == "yearly":
        s = date(ref.year-1, 1, 1); e = date(ref.year-1, 12, 31); label = str(s.year)
    else:
        raise ValueError("period must be daily|monthly|quarterly|yearly")
    return s, e, label

def build_report(period: str, today: date | None = None) -> dict:
    tx = load_tx()
    if today is None: today = datetime.now(TZ).date()
    s, e, label = _period_range(period, today)

    if tx.empty:
        df = pd.DataFrame(columns=tx.columns if not tx.empty else [])
    else:
        mask = (tx["date"] >= s) & (tx["date"] <= e)
        df = tx.loc[mask].copy()

    income_df  = df[df["amount_eur"] > 0] if not df.empty else df
    expense_df = df[df["amount_eur"] < 0] if not df.empty else df

    income = round(float(income_df["amount_eur"].sum()), 2) if not df.empty else 0.0
    spent  = round(float((-expense_df["amount_eur"]).sum()), 2) if not df.empty else 0.0
    net    = round(income - spent, 2)

    # Top-Kategorien (nur Ausgaben)
    cats = pd.DataFrame(columns=["category","spent"])
    IMG_DIR.mkdir(parents=True, exist_ok=True)
    images: list[Path] = []

    if not df.empty:
        if "category" not in df.columns: df["category"] = ""
        cat = expense_df.groupby("category", as_index=False)["amount_eur"].sum()
        cat["spent"] = -cat["amount_eur"]
        cats = cat[["category","spent"]].sort_values("spent", ascending=False)

        # Balken-Chart
        if not cats.empty:
            fig = plt.figure()
            top = cats.head(10).sort_values("spent")
            plt.barh(top["category"], top["spent"])
            plt.title(f"Top-Kategorien {period} {label}")
            plt.xlabel("EUR")
            p1 = IMG_DIR / f"{period}_{label}_categories.png"
            fig.tight_layout(); fig.savefig(p1); plt.close(fig)
            images.append(p1)

        # Kumulierte Ausgaben (>=Monat)
        if period in ("monthly","quarterly","yearly"):
            day = df.groupby("date", as_index=False)["amount_eur"].sum()
            day["cum_spent"] = (-day["amount_eur"].clip(upper=0)).cumsum()
            fig = plt.figure()
            plt.plot(day["date"], day["cum_spent"])
            plt.title(f"Kumulierte Ausgaben {period} {label}")
            plt.xlabel("Datum"); plt.ylabel("EUR")
            p2 = IMG_DIR / f"{period}_{label}_timeseries.png"
            fig.tight_layout(); fig.savefig(p2); plt.close(fig)
            images.append(p2)

    # Bewertung
    savings_rate = (income - spent) / income if income > 0 else 0.0
    rating = "✅ gut"
    if income == 0 and spent > 0: rating = "⚠️ nur Ausgaben"
    elif savings_rate < 0.05:     rating = "⚠️ sehr niedrig"
    elif savings_rate < 0.15:     rating = "🟡 mittel"

    return {
        "period": period, "label": label, "start": s, "end": e,
        "spent": spent, "income": income, "net": net,
        "top_categories": cats.to_dict("records"),
        "images": images,
        "rating": rating,
        "df": df
    }

# ------------------- HTML Rendering -------------------
def make_table_html(df: pd.DataFrame) -> str:
    if df is None or df.empty:
        return "<p>(Keine Buchungen im Zeitraum)</p>"
    cols = [c for c in ["date","account_id","payee","amount_eur","currency","category","tags","note"] if c in df.columns]
    tdf = df[cols].copy()
    tdf["date"] = pd.to_datetime(tdf["date"], errors="coerce").dt.strftime("%Y-%m-%d")
    if "amount_eur" in tdf:
        tdf["amount_eur"] = pd.to_numeric(tdf["amount_eur"], errors="coerce").fillna(0).round(2)
        tdf["amount_eur"] = tdf["amount_eur"].map(lambda x: f"{x:,.2f}".replace(",", " ").replace(".", ",").replace(" ", "."))
    html = tdf.to_html(index=False, escape=False)
    html = html.replace("<table", '<table style="border-collapse:collapse;width:100%;font-size:13px"') \
               .replace("<th", '<th style="border-bottom:1px solid #eaeaea;text-align:left;padding:6px 8px"') \
               .replace("<td", '<td style="border-bottom:1px solid #f4f4f4;padding:6px 8px"')
    return html

def render_html(rep: dict, table_html: str) -> str:
    name_html = f"<strong>{REPORT_NAME}</strong><br/>" if REPORT_NAME else ""
    addr_html = "<br/>".join(PROFILE_ADDR_LINES) if PROFILE_ADDR_LINES else ""

    cats = "".join([f"<li>{c['category']}: {c['spent']:.2f} €</li>" for c in rep["top_categories"][:10]]) or "<li>(keine)</li>"
    imgs = "".join([f'<p><small>{Path(p).name}</small> (Anhang)</p>' for p in rep["images"]])

    return f"""
    <div style="width:100%;font-family:-apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;font-size:14px;color:#222;">
      <div style="display:flex;justify-content:space-between;align-items:flex-start;gap:16px;">
        <div style="text-align:left;line-height:1.4">{name_html}{addr_html}</div>
        <div style="text-align:center;margin:0 auto;">
          <h2 style="margin:0 0 4px;">Finanz-Report</h2>
          <div style="font-weight:600">{rep['period'].capitalize()} – {rep['label']}</div>
          <div style="font-size:12px;color:#666">{rep['start']} bis {rep['end']}</div>
        </div>
        <div style="min-width:120px;"></div>
      </div>
      <hr style="margin:12px 0;border:none;border-top:1px solid #eee"/>
      <ul style="margin:0 0 8px 16px;padding:0;">
        <li><b>Ausgaben</b>: {rep['spent']:.2f} €</li>
        <li><b>Einnahmen</b>: {rep['income']:.2f} €</li>
        <li><b>Netto</b>: {rep['net']:.2f} €</li>
        <li><b>Bewertung</b>: {rep['rating']}</li>
      </ul>
      <p style="margin:12px 0 4px;"><b>Top-Kategorien (Ausgaben)</b></p>
      <ul style="margin:0 0 12px 16px;padding:0;">{cats}</ul>
      <p style="margin:12px 0 4px;"><b>Buchungen</b></p>
      <div style="overflow-x:auto;border:1px solid #eee;border-radius:6px;">
        {table_html}
      </div>
      <div style="margin-top:12px">{imgs}</div>
    </div>
    """

# ------------------- Farbiges PDF -------------------
def _kpi_card(ax, x0, y0, w, h, title, value, color):
    ax.add_patch(plt.Rectangle((x0, y0), w, h, facecolor="white", edgecolor=COLOR_BORDER, lw=0.8))
    ax.add_patch(plt.Rectangle((x0, y0+h-0.03), w, 0.03, facecolor=color, edgecolor=color, lw=0))
    ax.text(x0+0.02, y0+h-0.035, title, va="top", ha="left", fontsize=9, color=COLOR_MUTED)
    ax.text(x0+0.02, y0+0.04, value, va="bottom", ha="left", fontsize=13, color=COLOR_TEXT, fontweight="bold")

def _table_page(ax, df_page: pd.DataFrame, title: str):
    ax.axis("off")
    ax.text(0.0, 1.02, title, transform=ax.transAxes, fontsize=11, fontweight="bold", color=COLOR_TEXT)
    cols = [c for c in ["date","account_id","payee","amount_eur","currency","category","tags","note"] if c in df_page.columns]
    data = df_page[cols].copy()
    if "amount_eur" in data:
        data["amount_eur"] = data["amount_eur"].map(_fmt_eur)

    # Zebra-Hintergrund
    y0 = 0.96
    row_h = 0.028
    for i in range(len(data)):
        color = COLOR_TABLE_EVN if i % 2 == 0 else COLOR_TABLE_ODD
        ax.add_patch(plt.Rectangle((0.0, y0 - (i+1)*row_h), 1.0, row_h, facecolor=color, edgecolor="none"))

    table = ax.table(cellText=data.values.tolist(), colLabels=cols, loc="upper left",
                     cellLoc='left', colLoc='left', bbox=[0, 0, 1, 0.95])
    table.auto_set_font_size(False)
    table.set_fontsize(8)
    table.scale(1, 1.25)

    # Styles
    for (r, c), cell in table.get_celld().items():
        cell.set_edgecolor(COLOR_BORDER)
        if r == 0:
            cell.set_facecolor(COLOR_PRIMARY)
            cell.set_text_props(color="white", fontweight="bold")
        else:
            # Betrag rechtsbündig
            if cols[c] == "amount_eur":
                cell._loc = 'right'

def write_pdf_report(path: Path, rep: dict, df: pd.DataFrame, images: list[Path],
                     name: str, addr_lines: list[str]):
    path.parent.mkdir(parents=True, exist_ok=True)
    df_print = df.copy() if df is not None else pd.DataFrame()
    if "amount_eur" in df_print:
        df_print["amount_eur"] = pd.to_numeric(df_print["amount_eur"], errors="coerce").fillna(0).round(2)
    ROWS_PER_PAGE = 22

    with PdfPages(path) as pdf:
        # Seite 1: Header + KPIs + erste Tabelle
        fig = plt.figure(figsize=(8.27, 11.69))  # A4
        gs = fig.add_gridspec(nrows=100, ncols=1)

        # Header-Balken
        ax_bar = fig.add_subplot(gs[0:6, 0]); ax_bar.axis("off")
        ax_bar.add_patch(plt.Rectangle((0, 0), 1, 1, color=COLOR_PRIMARY))
        # Name/Adresse
        ax_head = fig.add_subplot(gs[6:18, 0]); ax_head.axis("off")
        header_lines = []
        if name: header_lines.append(name)
        header_lines += addr_lines
        ax_head.text(0.02, 0.8, "\n".join(header_lines), ha="left", va="top", fontsize=10, color=COLOR_TEXT)
        ax_head.text(0.5, 0.9, "Finanz-Report", ha="center", va="top", fontsize=16, color=COLOR_TEXT, fontweight="bold")
        ax_head.text(0.5, 0.62, f"{rep['period'].capitalize()} – {rep['label']}", ha="center", va="top", fontsize=11, color=COLOR_TEXT)
        ax_head.text(0.5, 0.48, f"{rep['start']} bis {rep['end']}", ha="center", va="top", fontsize=9, color=COLOR_MUTED)

        # KPI-Karten
        ax_kpi = fig.add_subplot(gs[18:30, 0]); ax_kpi.axis("off")
        _kpi_card(ax_kpi, 0.02, 0.05, 0.30, 0.9, "Ausgaben",  _fmt_eur(-abs(-rep['spent'])), COLOR_PRIMARY)
        _kpi_card(ax_kpi, 0.35, 0.05, 0.30, 0.9, "Einnahmen", _fmt_eur(rep['income']), COLOR_ACCENT)
        _kpi_card(ax_kpi, 0.68, 0.05, 0.30, 0.9, "Netto",     _fmt_eur(rep['net']), COLOR_TEXT)

        # Tabelle (erste Seite)
        ax_tbl = fig.add_subplot(gs[30:96, 0])
        if df_print.empty:
            ax_tbl.axis("off")
            ax_tbl.text(0.5, 0.5, "(Keine Buchungen im Zeitraum)", ha="center", va="center", fontsize=11, color=COLOR_MUTED)
        else:
            first = next(_chunks(df_print, ROWS_PER_PAGE))
            _table_page(ax_tbl, first, "Buchungen")

        # Footer
        ax_foot = fig.add_subplot(gs[96:100, 0]); ax_foot.axis("off")
        ax_foot.text(0.5, 0.5, f"Erstellt am {datetime.now(TZ).strftime('%Y-%m-%d %H:%M')}  •  {REPORT_NAME}",
                     ha="center", va="center", fontsize=8, color=COLOR_MUTED)

        fig.tight_layout()
        pdf.savefig(fig); plt.close(fig)

        # Folgeseiten: restliche Tabelle
        if not df_print.empty:
            remaining = list(_chunks(df_print.iloc[ROWS_PER_PAGE:], ROWS_PER_PAGE))
            for page_idx, chunk in enumerate(remaining, start=2):
                fig = plt.figure(figsize=(8.27, 11.69))
                ax = fig.add_subplot(111)
                _table_page(ax, chunk, f"Buchungen (Seite {page_idx})")
                fig.tight_layout(); pdf.savefig(fig); plt.close(fig)

        # Diagramm-Seite
        if images:
            fig = plt.figure(figsize=(8.27, 11.69))
            gs = fig.add_gridspec(nrows=100, ncols=1)
            ax_t = fig.add_subplot(gs[0:8, 0]); ax_t.axis("off")
            ax_t.text(0.5, 0.5, "Diagramme", ha="center", va="center", fontsize=14, color=COLOR_TEXT, fontweight="bold")
            slots = [(8, 54), (54, 100)]
            for img_path, (r0, r1) in zip(images[:2], slots):
                try:
                    img = mpimg.imread(str(img_path))
                    ax = fig.add_subplot(gs[r0:r1, 0]); ax.axis("off")
                    ax.imshow(img); ax.set_title(Path(img_path).name, fontsize=9, pad=6, color=COLOR_MUTED)
                except Exception:
                    ax = fig.add_subplot(gs[r0:r1, 0]); ax.axis("off")
                    ax.text(0.5, 0.5, f"(Bild konnte nicht geladen werden: {Path(img_path).name})",
                            ha="center", va="center", fontsize=10, color=COLOR_MUTED)
            fig.tight_layout(); pdf.savefig(fig); plt.close(fig)

# ------------------- Mailversand -------------------
def send_email(subject: str, html: str, attachments: list[Path]) -> None:
    if not (SMTP_USER and SMTP_PASS and EMAIL_FROM and EMAIL_TO):
        raise RuntimeError("SMTP/Gmail Variablen fehlen (SMTP_USER/SMTP_PASS/EMAIL_FROM/EMAIL_TO).")

    if not html or not html.strip():
        html = "<p>(Kein HTML-Inhalt erzeugt)</p>"
    plain = re.sub(r"<[^>]+>", "", html).replace("&nbsp;", " ").strip() or "(Kein Inhalt)"

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
        try:
            with open(p, "rb") as f:
                part = MIMEApplication(f.read(), Name=Path(p).name)
            part["Content-Disposition"] = f'attachment; filename="{Path(p).name}"'
            msg.attach(part)
        except FileNotFoundError:
            print(f"[WARN] Anhang fehlt: {p}")

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
    IMG_DIR.mkdir(parents=True, exist_ok=True)
    today = datetime.now(TZ).date()

    if period == "auto":
        periods: list[str] = []
        if _daily_at_20_local(): periods.append("daily")
        if today.day == 1:
            periods.append("monthly")
            if today.month in (1, 4, 7, 10): periods.append("quarterly")
            if today.month == 1: periods.append("yearly")
    else:
        periods = [period]

    for per in periods:
        rep = build_report(per, today=today)
        df_period = rep["df"].copy()
        # Für Nicht-Daily: Begrenze Tabelle (Übersichtsgröße im HTML), PDF bekommt alles
        df_html = df_period if per == "daily" else df_period.sort_values(["date","amount_eur"], ascending=[True, False]).head(30)

        # HTML
        table_html = make_table_html(df_html) if not df_html.empty else "<p>(Keine Buchungen im Zeitraum)</p>"
        html = render_html(rep, table_html)

        subject = f"Finanzen – {per} – {rep['label']} (Ausgaben {rep['spent']:.2f} €)"
        atts: list[Path] = list(rep["images"])

        # CSV (Daily)
        if per == "daily":
            dayfile = REPORT_DIR / f"report_{rep['label']}_day.csv"
            df_period.to_csv(dayfile, index=False, encoding="utf-8-sig")
            atts.append(dayfile)

        # PDF-„Urkunde“
        pdf_path = REPORT_DIR / f"statement_{rep['label']}.pdf"
        write_pdf_report(
            pdf_path,
            rep=rep,
            df=df_period,
            images=list(rep["images"]),
            name=REPORT_NAME,
            addr_lines=PROFILE_ADDR_LINES
        )
        atts.append(pdf_path)

        print(f"[REPORT] {per} {rep['label']} | income={rep['income']:.2f} | spent={rep['spent']:.2f} | net={rep['net']:.2f} | atts={len(atts)}")
        print(f"[DEBUG] HTML length: {len(html)} | to={EMAIL_TO}")

        if send:
            send_email(subject, html, atts)
            print(f"[MAIL] gesendet an: {EMAIL_TO}")

# ------------------- CLI -------------------
if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="Finanz-Reports per E-Mail senden (HTML + PDF)")
    ap.add_argument("period", choices=["daily","monthly","quarterly","yearly","auto"], help="Zeitraum")
    ap.add_argument("--send", action="store_true", help="E-Mail wirklich senden (sonst nur erzeugen)")
    args = ap.parse_args()
    run(args.period, send=args.send)
