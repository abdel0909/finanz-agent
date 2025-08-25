import os
import smtplib
import ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime
import pandas as pd
import matplotlib.pyplot as plt

# Import aus deinem Haupt-Agent
from finanzen_agent import create_workbook_if_missing, load_tx, build_report

SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.gmail.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))
SMTP_USER = os.getenv("bouardjaa@gmail.com")
SMTP_PASS = os.getenv("zwqdwuyxdzydtaqu")
EMAIL_FROM = os.getenv("bouardjaa@gmail.com", SMTP_USER)
EMAIL_TO = os.getenv("bouardjaa@gmail.com", SMTP_USER)
LOCAL_TZ = os.getenv("LOCAL_TZ", "Europe/Berlin")


def send_mail(subject: str, body: str):
    msg = MIMEMultipart()
    msg["From"] = EMAIL_FROM
    msg["To"] = EMAIL_TO
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "html"))

    context = ssl.create_default_context()
    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        server.starttls(context=context)
        server.login(SMTP_USER, SMTP_PASS)
        server.sendmail(EMAIL_FROM, EMAIL_TO.split(","), msg.as_string())


def main(period="daily", send=True):
    # Safety: Stelle sicher, dass Budget.xlsx existiert
    if not os.path.exists("Budget.xlsx"):
        print("[INFO] Budget.xlsx fehlt oder ist beschädigt → neu anlegen …")
        create_workbook_if_missing()

    # Lade Transaktionen
    tx = load_tx()

    # Report für Zeitraum bauen
    today = datetime.now().date()
    rep = build_report(period, today=today)

    # Textnachricht zusammenstellen
    subject = f"Finanz-Report ({period}) – {today}"
    body = f"""
    <h2>Finanzbericht: {period.capitalize()}</h2>
    <p><b>Ausgaben:</b> {rep['spent']:.2f} €<br>
    <b>Einnahmen:</b> {rep['income']:.2f} €<br>
    <b>Netto:</b> {rep['net']:.2f} €<br>
    <b>Bewertung:</b> {rep['rating']}</p>
    """

    if send:
        send_mail(subject, body)
        print("[DONE] Mail verschickt an:", EMAIL_TO)
    else:
        print("[TEST] Report erstellt, Mail nicht verschickt.")
        print(body)


if __name__ == "__main__":
    main("daily", send=True)
