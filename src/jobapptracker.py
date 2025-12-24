import os.path
from pathlib import Path
import base64
from datetime import datetime
import re

import pandas as pd

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from bs4 import BeautifulSoup
from email.utils import parseaddr
import string

SCOPES = [
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/spreadsheets"
]
APPLICATION_KEYWORDS = [
    "applying",
    "application",
    "interest"
]


    
def authentication():
    creds = None
  
    project_root = Path(__file__).parent.parent 
    credentials_folder = project_root / "credentials"
    credentials_path = credentials_folder / "credentials.json"
    token_path = credentials_folder / "token.json"
    credentials_folder.mkdir(exist_ok=True)

    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(str(token_path), SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
            str(credentials_path), SCOPES
            )
            creds = flow.run_local_server(port=0)
        with open(token_path, "w") as token:
            token.write(creds.to_json())

    return creds
    
def get_gmail_service(creds):
    return build("gmail", "v1", credentials=creds)

def get_sheets_service(creds):
    return build("sheets", "v4", credentials=creds)


def search_application_emails(service):
    subject_query = " OR ".join(
        [f'subject:"{k}"' for k in APPLICATION_KEYWORDS]
    )

    inbox_query = f"label:inbox ({subject_query})"

    results = service.users().messages().list(
        userId="me",
        q=inbox_query,
        maxResults=50
    ).execute()

    return results.get("messages", [])

def extract_body_from_payload(payload):
    body = ""

    if "parts" in payload:
        for part in payload["parts"]:
            body += extract_body_from_payload(part)
    else:
        mime_type = payload.get("mimeType", "")
        data = payload.get("body", {}).get("data")

        if data:
            decoded = base64.urlsafe_b64decode(data).decode(
                "utf-8", errors="ignore"
            )

            if mime_type == "text/plain":
                body += decoded
            elif mime_type == "text/html":
                soup = BeautifulSoup(decoded, "html.parser")
                body += soup.get_text(separator=" ")

    return body


def get_message_text(service, msg_id):
    msg = service.users().messages().get(
        userId="me",
        id=msg_id,
        format="full"
    ).execute()

    headers = msg["payload"]["headers"]
    subject = sender = date = ""

    for h in headers:
        if h["name"] == "Subject":
            subject = h["value"]
        elif h["name"] == "From":
            sender = h["value"]
        elif h["name"] == "Date":
            date = h["value"]

    body = extract_body_from_payload(msg["payload"])

    return subject, sender, date, body

def is_invalid_company(company: str) -> bool:
    if not company:
        return True

    if company.lower() == "this":
        return True

    if company.lower() == "the":
        return True
    
    if re.search(r"\d", company):
        return True

    return False

def extract_after_keyword(text: str, keyword: str) -> str | None:
    if not text or not keyword:
        return None

    pattern = rf"\b{re.escape(keyword)}\s+(\S+)"
    match = re.search(pattern, text, flags=re.IGNORECASE)

    if not match or is_invalid_company(match.group(1)):
        return None

    token = match.group(1)
    token = token.rstrip(string.punctuation)
    return token



def extract_company(from_header: str, body_text: str, subject: str) -> str:

    name, email_addr = parseaddr(from_header)
    email_addr = email_addr.lower()

    if email_addr.endswith("@myworkday.com"):
        company = email_addr.split("@")[0]
        if company:
            return company

    company = extract_after_keyword(body_text, "at")
    if company:
        return company

    company = extract_after_keyword(subject, "at")
    if company:
        return company
    
    
    return "Unknown"


def main():


    service = authentication()
    messages = search_application_emails(service)
    for msg in messages:
        subject, sender, date, body = get_message_text(service, msg["id"])
        company = extract_company(sender, subject, body)
        print(company)



if __name__ == "__main__":
  main()