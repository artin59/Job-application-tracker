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


SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]

APPLICATION_KEYWORDS = [
    "thank you for applying",
    "application received",
    "we received your application",
    "thanks for your interest",
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

    return build("gmail", "v1", credentials=creds)
    

def search_application_emails(service):
    query = " OR ".join([f'"{k}"' for k in APPLICATION_KEYWORDS])
    results = service.users().messages().list(
        userId="me",
        q=query,
        maxResults=50
    ).execute()

    return results.get("messages", [])

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

    body = ""

    parts = msg["payload"].get("parts", [])
    for part in parts:
        if part["mimeType"] == "text/plain":
            body = base64.urlsafe_b64decode(
                part["body"]["data"]
            ).decode("utf-8", errors="ignore")

    return subject, sender, date, body

def main():
    project_root = Path(__file__).parent.parent
    data_folder = project_root / "data"
    data_folder.mkdir(exist_ok=True)
    excel_path = data_folder / "applications.xlsx"

    service = authentication()
    messages = search_application_emails(service)
    for msg in messages:
        print(msg)

    for msg in messages:
        subject, sender, date, body = get_message_text(service, msg["id"])
        print(subject)



if __name__ == "__main__":
  main()