import os
import base64
from email.message import EmailMessage
from datetime import date
from typing import List

from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

SCOPES = ["https://www.googleapis.com/auth/gmail.send"]
GMAIL_TOKEN = "token_gmail.json"

def send_owner_snapshot_email(
    pdf_paths: List[str],
    report_day: date,
    data_start: date,
    data_end: date,
    to_email: str = "anthony@buzzcannabis.com",
):
    """
    Sends Owner Snapshot PDFs via Gmail API using JSON token (cron-safe).
    """

    if not os.path.exists(GMAIL_TOKEN):
        raise RuntimeError("❌ token_gmail.json not found — run Gmail auth first")

    creds = Credentials.from_authorized_user_file(GMAIL_TOKEN, SCOPES)

    if creds.expired and creds.refresh_token:
        creds.refresh(Request())
        with open(GMAIL_TOKEN, "w") as f:
            f.write(creds.to_json())

    service = build("gmail", "v1", credentials=creds)

    # ---------- Email ----------
    subject = f"Buzz Cannabis Owner Snapshot — {report_day.isoformat()}"

    body = f"""
Buzz Cannabis — Owner Snapshot

📅 Report Day:
{report_day.strftime('%A, %B %d, %Y')}

📊 Data Window:
{data_start.isoformat()} → {data_end.isoformat()}

Attached:
• Store Owner Snapshot PDFs
• All Stores Summary

This email was generated automatically.
"""

    msg = EmailMessage()
    msg["To"] = to_email
    msg["From"] = "me"
    msg["Subject"] = subject
    msg.set_content(body)

    # Attach PDFs
    for path in pdf_paths:
        if not os.path.exists(path):
            continue

        with open(path, "rb") as f:
            data = f.read()

        msg.add_attachment(
            data,
            maintype="application",
            subtype="pdf",
            filename=os.path.basename(path),
        )

    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()

    service.users().messages().send(
        userId="me",
        body={"raw": raw},
    ).execute()

    print(f"📧 Owner Snapshot emailed to {to_email}")
