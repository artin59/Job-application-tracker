import os.path
from pathlib import Path
import base64

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]

    
    
def main():
   
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

    service = build("gmail", "v1", credentials=creds)


if __name__ == "__main__":
  main()