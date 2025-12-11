import os
import pandas as pd
from datetime import datetime, timedelta
import smtplib, ssl
from email.message import EmailMessage

# ----------------- CONFIG -----------------
EXCEL_URL = os.getenv("EXCEL_URL")
EMAIL_USER = os.getenv("EMAIL_USER")
EMAIL_PASS = os.getenv("EMAIL_PASS")

# Multiple recipients
EMAIL_RECIPIENTS = [
    "haseebahmed2624@gmail.com"
]

# ----------------- TESTING -----------------
# Manual date for testing (optional)
target_date = pd.to_datetime("2025-03-12").date()  # 12 March 2025

# Columns
COL_MEMBER = "Member Name"
COL_DATE = "Date"
COL_STATUS = "Status"

# ------------------------------------------

def load_excel(url):
    df = pd.read_excel(url, engine='openpyxl')
    df.columns = [c.strip() for c in df.columns]
    return df

def send_email(subject, body):
    msg = EmailMessage()
    msg['From'] = EMAIL_USER
    msg['To'] = ", ".join(EMAIL_RECIPIENTS)
    msg['Subject'] = subject
    msg.set_content(body)

    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as smtp:
        smtp.login(EMAIL_USER, EMAIL_PASS)
        smtp.send_message(msg)

    print("Email sent to:", ", ".join(EMAIL_RECIPIENTS))

def main():
    df = load_excel(EXCEL_URL)
    df[COL_DATE] = pd.to_datetime(df[COL_DATE]).dt.date
    todays_data = df[df[COL_DATE] == target_date]

    all_members = df[COL_MEMBER].unique() if not df.empty else []
    submitted_members = todays_data[COL_MEMBER].unique() if not todays_data.empty else []
    missing_members = [m for m in all_members if m not in submitted_members]

    if missing_members:
        lines = [
            f"Daily Status Report Check — Missing Reports for {target_date}",
            "",
            "The following members did NOT submit their daily report:",
            ""
        ]
        for m in missing_members:
            lines.append(f"- {m}")
        email_body = "\n".join(lines)
        send_email(subject=f"Missing Status Reports — {target_date}", body=email_body)
    else:
        print("All members submitted their status report or no data for this date.")

if __name__ == "__main__":
    main()
