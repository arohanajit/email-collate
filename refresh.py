import os
import base64
from datetime import datetime, date
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Import the PDF processing function
from a import process_and_update

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

def get_gmail_service():
    print("Authenticating with Gmail...")
    creds = None

    # Remove the existing token.json file
    if os.path.exists('token.json'):
        print("Removing existing token.json file.")
        os.remove('token.json')
    
    if not creds or not creds.valid:
        print("Starting new authentication flow...")
        flow = InstalledAppFlow.from_client_secrets_file(
            'd_credentials.json', SCOPES)
        creds = flow.run_local_server(port=0)
        print("Saving new credentials to token.json...")
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    print("Authentication successful.")
    return build('gmail', 'v1', credentials=creds)