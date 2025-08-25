#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import os
import pandas as pd
import matplotlib.pyplot as plt
from finanzen_agent import create_workbook_if_missing
from __future__ import annotations
import os, io, json, smtplib, ssl, base64
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from pathlib import Path
from datetime import datetime, date, timedelta
from zoneinfo import ZoneInfo

import pandas as pd
import matplotlib.pyplot as plt

BASE = Path(__file__).parent.resolve()
WORKBOOK = BASE / "Budget.xlsx"
REPORT_DIR = BASE / "reports"

TZ = ZoneInfo(os.getenv("LOCAL_TZ", "Europe/Berlin"))

def _load_tx() -> pd.DataFrame:
    if not WORKBOOK.exists():
        raise FileNotFoundError("Budget.xlsx nicht gefunden – zuerst finanzen_agent.py laufen lassen.")
    tx = pd.read_excel(WORKBOOK, sheet_name="Transactions", engine="openpyxl")
    if tx.empty:
        return tx
    tx["date"] = pd.to_datetime(tx["date"], errors="coerce").dt.date
    tx["amount_eur"] = pd.to_numeric(tx["amount_eur"], errors="coerce").fillna(0.0)
    tx["category"] = tx.get("category", "").astype(str)
    return tx

def _sum_spent(df: pd.DataFrame) -> float:
    # Ausgaben sind negativ -> Betrag (absolut) der <0 Summen
    return float((-df.loc[df["amount_eur"] < 0, "amount_eur"]).sum())

def _period_range(period: str, ref: date) -> tuple[date, date, str]:
    """Gibt (start, end_inclusive, label) für vergangene Perioden zurück."""
    if period == "daily":
        s = ref
        e = ref
        label = ref.strftime("%Y-%m-%d")
    elif period == "monthly":
        # kompletter Vormonat
        first_this = ref.replace(day=1)
        last_prev = first_this - timedelta(days=1)
        s = last_prev.replace(day=1)
        e = last_prev
        label = s.strftime("%Y-%m")
    elif period == "quarterly":
        m = ((ref.month - 1) // 3) * 3 + 1   # aktuelle Q-Start
        first_this_q = date(ref.year, m, 1)
        last_prev = first_this_q - timedelta(days=1)
        m0 = ((last_prev.month - 1)//3)*3 + 1
        s = date(last_prev.year, m0, 1)
        # Quartalsende
        if m0 in (1,4,7): e = date(last_prev.year, m0+2, (date(last_prev.year, m0+3,1) - timedelta(days=1)).day)
        else: e = last_prev
        label = f"{s.strftime('%Y-Q')}{(s.month-1)//3 + 1}"
    elif period == "yearly":
        s = date(ref.year-1, 1, 1)
        e = date(ref.year-1, 12, 31)
        label = str(s.year)
    else:
        raise ValueError("period must be one of daily, monthly, quarterly, yearly")
    return s, e, label

def _chart_path(name: str) -> Path:
    (REPORT_DIR / "imgs").mkdir(parents=True, exist_ok=True)
    return REPORT_DIR / "imgs" / name

def build_report(period: str, today: date | None = None) -> dict:
    """Berechnet Summen, erzeugt Diagramme, liefert Dict für E-Mail."""
    tx = _load_tx()
    if today is None:
        today = datetime.now(TZ).date()

    s, e, label = _period_range(period, today)
    mask = (tx["date"] >= s) & (tx["date"] <= e)
    df = tx.loc[mask].copy()

    # Summen
    total_spent = _sum_spent(df)
    income = float(df.loc[df["amount_eur"] > 0, "amount_eur"].sum())
    net = float(df["amount_eur"].sum())

    # Kategorie-Summen (nur Ausgaben)
    cat = df[df["amount_eur"] < 0].copy()
    cat = cat.groupby("category", as_index=False)["amount_eur"].sum().sort_values("amount_eur")
    cat["spent_abs"] = -cat["amount_eur"]

    # Diagramme
    images = []

    if not cat.empty:
        # Balken: Top-Kategorien (Ausgaben)
        fig = plt.figure()
        top = cat.sort_values("spent_abs", ascending=True).tail(10)
        plt.barh(top["category"], top["spent_abs"])
        plt.title(f"Top-Kategorien {period} {label}")
        plt.xlabel("EUR")
        p = _chart_path(f"{period}_{label}_categories.png")
        fig.tight_layout()
        fig.savefig(p)
        plt.close(fig)
        images.append(p)

    # Zeitreihe: tägliche Summe (nur bei month/quarter/year sinnvoll)
    if period in ("monthly","quarterly","yearly") and not df.empty:
        fig = plt.figure()
        day = df.groupby("date", as_index=False)["amount_eur"].sum()
        day["cum_spent"] = (-day["amount_eur"].clip(upper=0)).cumsum()
        plt.plot(day["date"], day["cum_spent"])
        plt.title(f"Kumulierte Ausgaben {period} {label}")
        plt.xlabel("Datum"); plt.ylabel("EUR")
        p = _chart_path(f"{period}_{label}_timeseries.png")
        fig.tight_layout()
        fig.savefig(p)
        plt.close(fig)
        images.append(p)

    # Performance / Bewertung (einfaches Scoring)
    # Savings Rate über die Periode
    saving_rate = (income + min(0.0, -total_spent) + net)  # redundant – wir wählen simpler:
    expenses_abs = total_spent
    savings_rate = (income + net) / income if income > 0 else 0.0  # (Income - |Expenses|)/Income ~ näherungsweise
    rating = "✅ gut"
    if income == 0 and expenses_abs > 0:
        rating = "⚠️ nur Ausgaben"
    elif savings_rate < 0.05:
        rating = "⚠️ sehr niedrig"
    elif savings_rate < 0.15:
        rating = "🟡 mittel"

    return {
        "period": period,
        "label": label,
        "start": s, "end": e,
        "spent": round(total_spent, 2),
        "income": round(income, 2),
        "net": round(net, 2),
        "top_categories": cat[["category","spent_abs"]].rename(columns={"spent_abs":"spent"}).to_dict("records"),
        "images": images,
        "rating": rating,
    }

# ----------------------- email -----------------------
def _send_smtp(subject: str, html: str, attachments: list[Path], to_addrs: list[str]):
    host = os.getenv("SMTP_SERVER", "smtp-mail.outlook.com")
    port = int(os.getenv("SMTP_PORT", "587"))
    user = os.getenv("SMTP_USER")
    pwd = os.getenv("SMTP_PASS")
    sender = os.getenv("EMAIL_FROM", user)

    if not user or not pwd:
        raise RuntimeError("SMTP_USER/SMTP_PASS nicht gesetzt (GitHub Secret).")

    msg = MIMEMultipart("related")
    msg["Subject"] = subject
    msg["From"] = sender
    msg["To"] = ", ".join(to_addrs)

    alt = MIMEMultipart("alternative")
    alt.attach(MIMEText(html, "html", "utf-8"))
    msg.attach(alt)

    for p in attachments:
        with open(p, "rb") as f:
            part = MIMEApplication(f.read(), Name=p.name)
        part["Content-Disposition"] = f'attachment; filename="{p.name}"'
        msg.attach(part)

    ctx = ssl.create_default_context()
    with smtplib.SMTP(host, port) as s:
        s.starttls(context=ctx)
        s.login(user, pwd)
        s.sendmail(sender, to_addrs, msg.as_string())

def _send_graph(subject: str, html: str, attachments: list[Path], to_addrs: list[str]):
    """
    Microsoft Graph (Client-Credentials): braucht TENANT_ID, CLIENT_ID, CLIENT_SECRET, GRAPH_SENDER
    """
    import requests
    tenant = os.getenv("AZURE_TENANT_ID")
    client = os.getenv("AZURE_CLIENT_ID")
    secret = os.getenv("AZURE_CLIENT_SECRET")
    sender = os.getenv("GRAPH_SENDER")  # UPN oder mailbox (z. B. me@domain.com)

    if not (tenant and client and secret and sender):
        raise RuntimeError("Graph-Secrets fehlen (AZURE_TENANT_ID/CLIENT_ID/CLIENT_SECRET/GRAPH_SENDER).")

    # Token holen
    token_url = f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token"
    data = {"client_id": client, "client_secret": secret, "scope": "https://graph.microsoft.com/.default", "grant_type": "client_credentials"}
    tok = requests.post(token_url, data=data).json()["access_token"]

    # Attachments als base64
    atts = []
    for p in attachments:
        b64 = base64.b64encode(p.read_bytes()).decode("utf-8")
        atts.append({
            "@odata.type": "#microsoft.graph.fileAttachment",
            "name": p.name,
            "contentBytes": b64
        })

    body = {
        "message": {
            "subject": subject,
            "body": {"contentType": "HTML", "content": html},
            "toRecipients": [{"emailAddress": {"address": addr}} for addr in to_addrs],
            "attachments": atts
        },
        "saveToSentItems": "true"
    }
    url = f"https://graph.microsoft.com/v1.0/users/{sender}/sendMail"
    r = requests.post(url, headers={"Authorization": f"Bearer {tok}", "Content-Type": "application/json"}, data=json.dumps(body))
    if r.status_code >= 300:
        raise RuntimeError(f"Graph sendMail failed: {r.status_code} {r.text}")

def send_email(subject: str, html: str, attachments: list[Path], to_addrs: list[str]):
    # wenn Graph-Umgebung vorhanden, Graph nutzen – sonst SMTP
    try:
        if os.getenv("AZURE_TENANT_ID"):
            _send_graph(subject, html, attachments, to_addrs)
        else:
            _send_smtp(subject, html, attachments, to_addrs)
    except Exception as e:
        raise

def render_html(rep: dict) -> str:
    cats = "".join([f"<li>{c['category']}: {c['spent']:.2f} €</li>" for c in rep["top_categories"][:10]])
    imgs = "".join([f'<p><b>{p.name}</b><br><i>Angehängt</i></p>' for p in rep["images"]])
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
    <ul>{cats or "<li>(keine)</li>"}</ul>
    <hr />{imgs}
    """

def maybe_send_daily_at_20() -> bool:
    # Schutz, falls Workflow sowohl 18:00 als auch 19:00 UTC läuft (DST): nur um 20:00 lokaler Zeit senden
    now = datetime.now(TZ)
    return now.hour == 20

def main():
    
       # Safety: Stelle sicher, dass Budget.xlsx existiert
    if not os.path.exists("Budget.xlsx"):
        print("[INFO] Budget.xlsx fehlt oder ist beschädigt → neu anlegen …")
        create_workbook_if_missing()
    import argparse as _ap
    p = _ap.ArgumentParser(description="Berichtsversand per Mail")
    p.add_argument("period", choices=["daily","monthly","quarterly","yearly","auto"], help="Welche Periode?")
    p.add_argument("--send", action="store_true", help="E-Mail wirklich senden (sonst nur Charts/Reports erzeugen)")
    p.add_argument("--to", default=os.getenv("EMAIL_TO",""), help="Empfänger, Komma-getrennt (oder Secret EMAIL_TO)")
    args = p.parse_args()

    periods = []
    today = datetime.now(TZ).date()

    if args.period == "auto":
        # immer Tagesreport (20:00), Monatsreport am 1., Quartal am 1. in Jan/Apr/Jul/Okt, Jahresreport am 1. Jan
        if maybe_send_daily_at_20():
            periods.append("daily")
        if today.day == 1:
            periods.append("monthly")
            if today.month in (1,4,7,10):
                periods.append("quarterly")
            if today.month == 1:
                periods.append("yearly")
    else:
        if args.period == "daily" and not maybe_send_daily_at_20():
            # nicht 20 Uhr? trotzdem erzeugen, aber nicht senden (oder mit --send erzwingen)
            pass
        periods.append(args.period)

    to_list = [x.strip() for x in (args.to or "").split(",") if x.strip()]
    REPORT_DIR.mkdir(parents=True, exist_ok=True)

    for per in periods:
        rep = build_report(per, today=today)
        subject = f"Finanzen – {per} – {rep['label']} (Ausgaben {rep['spent']:.2f} €)"
        html = render_html(rep)
        attachments = rep["images"]

        # CSV der Periode beilegen (nice to have)
        # (Aus der Reports-Mappe kommt bereits report_YYYY-MM.csv – hier nur der Komfort)
        if per == "monthly":
            p = REPORT_DIR / f"report_{rep['label']}.csv"
            if p.exists(): attachments = attachments + [p]
        elif per == "daily":
            # Tages-CSV erzeugen
            tx = _load_tx()
            mask = (tx["date"] >= rep["start"]) & (tx["date"] <= rep["end"])
            dayfile = REPORT_DIR / f"report_{rep['label']}_day.csv"
            tx.loc[mask].to_csv(dayfile, index=False)
            attachments = attachments + [dayfile]

        print(f"[REPORT] {per} {rep['label']} | spent={rep['spent']:.2f} | attachments={len(attachments)}")

        if args.send and to_list:
            send_email(subject, html, attachments, to_list)
            print(f"[MAIL] gesendet an: {to_list}")

if __name__ == "__main__":
    main()
