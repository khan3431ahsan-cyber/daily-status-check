import os
import pandas as pd
from datetime import datetime, timedelta
import smtplib, ssl
from email.message import EmailMessage
import requests
from io import BytesIO

# ----------------- CONFIG -----------------

EXCEL_URL = os.getenv("EXCEL_URL")  # Online Excel/CSV URL
EMAIL_USER = os.getenv("EMAIL_USER")      # Sender email (Gmail SMTP)
EMAIL_PASS = os.getenv("EMAIL_PASS")      # Sender email password / App Password
EMAIL_RECIPIENT = "haseebahmed2624@gmail.com"  # Fixed recipient

# Default: check yesterday’s date
target_date = (datetime.utcnow() - timedelta(days=1)).date()

# Required columns
COL_MEMBER = "Member Name"
COL_DATE = "Date"
COL_STATUS = "Status"

# ------------------------------------------

def load_excel(url):
    """Load Excel or CSV from URL safely"""
    if url.endswith(".csv"):
        df = pd.read_csv(url)
    else:
        # For xlsx, download file first and read with openpyxl engine
        resp = requests.get(url)
        resp.raise_for_status()  # Fail if request fails
        df = pd.read_excel(BytesIO(resp.content), engine='openpyxl')

    # Clean column names
    df.columns = [c.strip() for c in df.columns]
    return df


def send_email(subject, body):
    """Send email via Gmail SMTP"""
    msg = EmailMessage()
    msg['From'] = EMAIL_USER
    msg['To'] = EMAIL_RECIPIENT
    msg['Subject'] = subject
    msg.set_content(body)

    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as smtp:
        smtp.login(EMAIL_USER, EMAIL_PASS)
        smtp.send_message(msg)

    print("Email sent to admin:", EMAIL_RECIPIENT)


def main():
    df = load_excel(EXCEL_URL)

    # Format the date column
    df[COL_DATE] = pd.to_datetime(df[COL_DATE]).dt.date

    # Filter for target date
    todays_data = df[df[COL_DATE] == target_date]

    # List of all unique employees
    all_members = df[COL_MEMBER].unique()

    # List of members who submitted status report
    submitted_members = todays_data[COL_MEMBER].unique()

    # Members who did NOT submit
    missing_members = [m for m in all_members if m not in submitted_members]

    if missing_members:
        # Prepare email body
        lines = [
            f"Daily Status Report Check — Missing Reports for {target_date}",
            "",
            "The following members did NOT submit their daily report:",
            ""
        ]

        for member in missing_members:
            lines.append(f"- {member}")

        email_body = "\n".join(lines)

        send_email(
            subject=f"Missing Status Reports — {target_date}",
            body=email_body
        )

    else:
        print("All members submitted their status report.")


if __name__ == "__main__":
    main()
