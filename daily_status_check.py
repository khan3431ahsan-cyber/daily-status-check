import os
import sys
import argparse
from datetime import datetime, timedelta
import pandas as pd
from email.message import EmailMessage
import smtplib
from dotenv import load_dotenv
import re

load_dotenv()

SMTP_HOST = os.getenv("SMTP_HOST")
SMTP_PORT = int(os.getenv("SMTP_PORT", "465"))
SMTP_USER = os.getenv("SMTP_USER")
SMTP_PASS = os.getenv("SMTP_PASS")
FROM_EMAIL = os.getenv("FROM_EMAIL", SMTP_USER)
HASEEB_EMAIL = os.getenv("HASEEB_EMAIL", "haseebahmed2624@gmail.com")

def find_column(df, keywords):
    lower_cols = {c.lower(): c for c in df.columns}
    for kw in keywords:
        for col_lower, col_orig in lower_cols.items():
            if kw in col_lower:
                return col_orig
    return None

def load_file(path):
    if path.lower().endswith(('.xls', '.xlsx')):
        df = pd.read_excel(path, engine='openpyxl')
    elif path.lower().endswith('.csv'):
        df = pd.read_csv(path)
    else:
        raise ValueError("Unsupported file type. Provide .xlsx, .xls or .csv")
    return df

def normalize_date_col(df, date_col):

    s = pd.to_datetime(df[date_col], dayfirst=True, errors='coerce')
    return s

def send_email(subject, body, to_addrs, dry_run=False):
    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = FROM_EMAIL
    msg['To'] = to_addrs if isinstance(to_addrs, str) else ", ".join(to_addrs)
    msg.set_content(body)

    if dry_run:
        print("=== DRY RUN: EMAIL ===")
        print("To:", msg['To'])
        print("Subject:", subject)
        print(body)
        print("=====================")
        return True

    if not all([SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASS, FROM_EMAIL]):
        raise RuntimeError("SMTP credentials not fully configured in environment variables.")

    if SMTP_PORT == 465:
        with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT) as smtp:
            smtp.login(SMTP_USER, SMTP_PASS)
            smtp.send_message(msg)
    else:
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as smtp:
            smtp.ehlo()
            smtp.starttls()
            smtp.login(SMTP_USER, SMTP_PASS)
            smtp.send_message(msg)
    return True

def is_blank(x):
    if pd.isna(x):
        return True
    s = str(x).strip()
    return s == "" or s.lower() in ["nan", "none", "n/a", "-"]

def main(args):
    path = args.file
    target_date = args.date
    dry_run = args.dry_run

    if target_date:

        target = datetime.strptime(target_date, "%Y-%m-%d").date()
    else:
        target = (datetime.now() - timedelta(days=1)).date()

    print(f"Loading file: {path}")
    df = load_file(path)
    if df.shape[0] == 0:
        print("Input file has no rows. Exiting.")
        return

    name_col = find_column(df, ['member', 'name', 'member name'])
    date_col = find_column(df, ['date', 'timestamp'])
    status_col = find_column(df, ['status', 'status report', 'report'])
    hours_col = find_column(df, ['working', 'hours', 'working hours'])
    email_col = find_column(df, ['email', 'email address', 'e-mail'])

    if not name_col or not date_col:
        print("Could not detect required columns. Found columns:", df.columns.tolist())
        raise SystemExit("Required columns missing: Member Name and Date are mandatory.")

    print("Detected columns:")
    print("  Name column:", name_col)
    print("  Date column:", date_col)
    print("  Status column:", status_col)
    print("  Hours column:", hours_col)
    print("  Email column:", email_col)

    # Normalize date column
    df['_parsed_date'] = normalize_date_col(df, date_col)
    df['_parsed_date_only'] = df['_parsed_date'].dt.date

    # Build set of all known employees 
    all_employees = sorted(df[name_col].dropna().astype(str).str.strip().unique())

    # Filter rows for target date
    df_target = df[df['_parsed_date_only'] == target]

    # Find employees who have at least one non-empty status report row for target date
    present = set()
    present_info = {}  
    for idx, row in df_target.iterrows():
        nm = str(row[name_col]).strip()
        status_val = row[status_col] if status_col and status_col in row else None
        hours_val = row[hours_col] if hours_col and hours_col in row else None
        email_val = row[email_col] if email_col and email_col in row else None
        # if status exists and not blank, count as present
        if status_col and not is_blank(status_val):
            present.add(nm)
            present_info[nm] = {
                "status": str(status_val).strip(),
                "hours": hours_val,
                "email": email_val,
                "timestamp": row.get(date_col)
            }
        else:
            # If status blank but hours present and >0, optionally treat as present (configurable)
            if hours_col and not is_blank(hours_val):
                try:
                
                    h = float(hours_val)
                    if h > 0:
                        present.add(nm)
                        present_info[nm] = {
                            "status": str(status_val).strip() if status_val is not None else "",
                            "hours": hours_val,
                            "email": email_val,
                            "timestamp": row.get(date_col)
                        }
                except Exception:
                    # if not numeric but non-blank, still treat as present
                    present.add(nm)
                    present_info[nm] = {
                        "status": str(status_val).strip() if status_val is not None else "",
                        "hours": hours_val,
                        "email": email_val,
                        "timestamp": row.get(date_col)
                    }

    missing = []
    for emp in all_employees:
        if emp not in present:
            # find last known email or hours for reporting
            emp_rows = df[df[name_col].astype(str).str.strip() == emp]
            # pick any row to get email/hours if present
            sample = emp_rows.iloc[0] if len(emp_rows) > 0 else None
            sample_email = sample[email_col] if (sample is not None and email_col in sample) else None
            sample_hours = sample[hours_col] if (sample is not None and hours_col in sample) else None
            missing.append({
                "name": emp,
                "email": sample_email if not is_blank(sample_email) else "",
                "hours": sample_hours if not is_blank(sample_hours) else ""
            })

    # Compose and send summary to Haseeb if anyone missing
    if missing:
        body_lines = [
            f"Daily status check for {target.isoformat()}",
            "",
            "The following employees did NOT submit a status report (or report was blank) on that date:",
            ""
        ]
        for m in missing:
            body_lines.append(f"- {m['name']}\tEmail: {m['email']}\tWorkingHours (sample): {m['hours']}")
        body_lines.append("")
        body_lines.append("This email was autogenerated by the daily_status_check script.")
        subject = f"[Status Alert] Missing status reports for {target.isoformat()}"
        body = "\n".join(body_lines)
        print("Sending missing-report summary to Haseeb:", HASEEB_EMAIL)
        send_email(subject, body, HASEEB_EMAIL, dry_run=dry_run)
    else:
        print("All employees present for", target.isoformat())

    #SPECIAL CHECK FOR HASEEB (SIMPLIFIED)
    haseeb_rows = df_target[df_target[name_col].astype(str).str.strip().str.lower() == "haseeb"]

    # Case 1: No row found for Haseeb → fully absent
    if len(haseeb_rows) == 0:
        subject = f"You were absent on {target.isoformat()}"
        body = (
            f"Hello Haseeb,\n\n"
            f"You are absent that day and you don't do any work on this day.\n\n"
            "Regards,\nSharkstack"
        )
        print("Haseeb absent (no entry found) → sending email.")
        send_email(subject, body, HASEEB_EMAIL, dry_run=dry_run)

    else:
        # Take any row for Haseeb
        row = haseeb_rows.iloc[0]
        h_hours = row[hours_col] if hours_col and hours_col in row else None

        # Condition: working hours missing or 0 or blank
        if is_blank(h_hours) or str(h_hours).strip() in ["0", "0.0"]:
            subject = f"You were absent on {target.isoformat()}"
            body = (
                f"Hello Haseeb,\n\n"
                f"You are absent that day and you don't do any work on this day.\n\n"
                "Regards,\nAutomated Attendance Bot"
            )
            print("Haseeb hours empty/zero → sending absent email.")
            send_email(subject, body, HASEEB_EMAIL, dry_run=dry_run)
        else:
            print(f"Haseeb present with working hours: {h_hours}")



if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Daily status check script")
    parser.add_argument("file", help="Path to Excel (.xlsx/.xls) or CSV file containing status reports")
    parser.add_argument("--date", help="target date in YYYY-MM-DD (default=yesterday)", default=None)
    parser.add_argument("--dry-run", help="Do not send emails, only print what would be sent", action="store_true")
    args = parser.parse_args()
    main(args)
