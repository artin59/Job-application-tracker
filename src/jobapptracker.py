import os.path
from pathlib import Path
import base64
from datetime import datetime
import re
from dotenv import load_dotenv
import pandas as pd

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from bs4 import BeautifulSoup
from email.utils import parseaddr
import string


project_root = Path(__file__).parent.parent
env_path = project_root / "credentials" / ".env"
load_dotenv(dotenv_path=env_path)

SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")
START_ROW = 84

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

def format_date(date_string: str) -> str:
    date_string = re.sub(r"\s\([A-Za-z]+\)$", "", date_string)
    original_format = "%a, %d %b %Y %H:%M:%S %z"  
    dt_object = datetime.strptime(date_string, original_format)
    formatted_date = dt_object.strftime("%B %d, %Y")  
    return formatted_date
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

def get_next_empty_row(sheets_service, spreadsheet_id):
    try:
        for chunk_start in range(START_ROW, START_ROW + 1000, 100):  
            chunk_end = chunk_start + 99
            range_name = f"A{chunk_start}:F{chunk_end}"
            
            result = sheets_service.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=range_name
            ).execute()
            
            values = result.get('values', [])
            
            for i, row in enumerate(values):
                row_num = chunk_start + i
                if not row or all(cell == "" for cell in row[:1]):  #
                    return row_num
            
            if len(values) < 100:
                return chunk_start + len(values)
        
        return START_ROW + 1000
        
    except HttpError as error:
        print(f"Error finding next empty row: {error}")
        return START_ROW 
    
def copy_dropdown_from_above(sheets_service, spreadsheet_id, start_row, num_rows):

    source_row_index = start_row - 2 

    requests = [
        {
            "copyPaste": {
                "source": {
                    "sheetId": 0,
                    "startRowIndex": source_row_index,
                    "endRowIndex": source_row_index + 1,
                    "startColumnIndex": 4, 
                    "endColumnIndex": 5
                },
                "destination": {
                    "sheetId": 0,
                    "startRowIndex": start_row - 1,
                    "endRowIndex": start_row - 1 + num_rows,
                    "startColumnIndex": 4,
                    "endColumnIndex": 5
                },
                "pasteType": "PASTE_DATA_VALIDATION"
            }
        }
    ]

    sheets_service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"requests": requests}
    ).execute()



def update_google_sheet(sheets_service, spreadsheet_id, new_entries):
    if not new_entries:
        return
    
    try:
        start_row = get_next_empty_row(sheets_service, spreadsheet_id)
        
        values = []
        for entry in new_entries:
            values.append([
                entry["Company Name"],  
                "",                     
                "",                     
                entry["Date Applied"],  
                "Applied ",              
                ""                      
            ])
        
        end_row = start_row + len(values) - 1
        range_name = f"A{start_row}:F{end_row}"
        
        body = {
            "values": values
        }
        
        update_result = sheets_service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=range_name,
            valueInputOption="USER_ENTERED",
            body=body
        ).execute()
        
        copy_dropdown_from_above(
            sheets_service,
            spreadsheet_id,
            start_row,
            len(values)
        )

        return start_row
        
    except HttpError as error:
        return None
    
def main():
    creds = authentication()
    gmail_service = get_gmail_service(creds)
    sheets_service = get_sheets_service(creds)

    messages = search_application_emails(gmail_service)

    new_entries = []

    for msg in reversed(messages):
        subject, sender, date, body = get_message_text(gmail_service, msg["id"])
        company = extract_company(sender, subject, body)
        if company != "Unknown":
            entry = {
                "Company Name": company.capitalize(),
                "Date Applied": format_date(date),
            }
            new_entries.append(entry)
    
    
    if new_entries:
        update_google_sheet(sheets_service, SPREADSHEET_ID, new_entries)     

if __name__ == "__main__":
  main()