import os
import pandas as pd
from datetime import datetime, timedelta
import smtplib, ssl
from email.message import EmailMessage

# ----------------- CONFIG -----------------

EXCEL_URL = os.getenv("EXCEL_URL")  # Direct XLSX export link
EMAIL_USER = os.getenv("EMAIL_USER")
EMAIL_PASS = os.getenv("EMAIL_PASS")
EMAIL_RECIPIENT = "haseebahmed2624@gmail.com"

# Check yesterday’s date by default
target_date = (datetime.utcnow() - timedelta(days=1)).date()

# Required columns
COL_MEMBER = "Member Name"
COL_DATE = "Date"
COL_STATUS = "Status"

# ------------------------------------------

def load_excel(url):
    """Load online Excel file directly using pandas + openpyxl"""
    df = pd.read_excel(url, engine='openpyxl')
    df.columns = [c.strip() for c in df.columns]  # clean column names
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

    # All members
    all_members = df[COL_MEMBER].unique()

    # Members who submitted
    submitted_members = todays_data[COL_MEMBER].unique()

    # Missing members
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
        print("All members submitted their status report.")


if __name__ == "__main__":
    main()
