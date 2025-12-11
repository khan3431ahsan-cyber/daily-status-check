import os
import pandas as pd
from datetime import datetime
import smtplib, ssl
from email.message import EmailMessage

# ----------------- CONFIG -----------------
# Load from environment variables
EXCEL_URL = os.getenv("EXCEL_URL")  # Online Excel URL (xlsx or csv)
EMAIL_USER = os.getenv("EMAIL_USER")  # Your email (SMTP sender)
EMAIL_PASS = os.getenv("EMAIL_PASS")  # App password / SMTP password
EMAIL_RECIPIENT = os.getenv("EMAIL_RECIPIENT")  # Where summary will go

# Optional: date to check (YYYY-MM-DD), default = yesterday
TARGET_DATE = os.getenv("TARGET_DATE")
if TARGET_DATE:
    target_date = datetime.strptime(TARGET_DATE, "%Y-%m-%d").date()
else:
    from datetime import timedelta
    target_date = (datetime.utcnow() - timedelta(days=1)).date()

# Columns expected in the Excel
COL_MEMBER = "Member Name"
COL_DATE = "Date"
COL_STATUS = "Status"

# ------------------------------------------

def load_excel(url):
    """Load Excel from URL (xlsx or csv)"""
    if url.endswith(".csv"):
        df = pd.read_csv(url)
    else:
        df = pd.read_excel(url)
    # Normalize column names
    df.columns = [c.strip() for c in df.columns]
    return df

def send_email(subject, body, recipient):
    msg = EmailMessage()
    msg['From'] = EMAIL_USER
    msg['To'] = recipient
    msg['Subject'] = subject
    msg.set_content(body)

    context = ssl.create_default_context()
    with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as smtp:
        smtp.login(EMAIL_USER, EMAIL_PASS)
        smtp.send_message(msg)
    print(f"Email sent to {recipient}")

def main():
    df = load_excel(EXCEL_URL)

    # Convert Date column to datetime.date
    df[COL_DATE] = pd.to_datetime(df[COL_DATE]).dt.date

    # Filter rows for the target date
    df_target = df[df[COL_DATE] == target_date]

    # All members
    all_members = df[COL_MEMBER].unique()

    # Members who submitted
    submitted_members = df_target[COL_MEMBER].unique()

    # Missing members
    missing_members = [m for m in all_members if m not in submitted_members]

    if missing_members:
        lines = [f"Missing status reports for {target_date}:", ""]
        for m in missing_members:
            # Optional: include other details if available
            lines.append(f"- {m}")
        body = "\n".join(lines)
        send_email(f"Missing Status Reports for {target_date}", body, EMAIL_RECIPIENT)
    else:
        print(f"All members submitted status for {target_date}")

if __name__ == "__main__":
    main()
