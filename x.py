# gmail_downloader.py

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

def get_attachments(service, user_id, msg_id, path1, path2, term1, term2):
    try:
        print(f"Fetching message with ID: {msg_id}")
        message = service.users().messages().get(userId=user_id, id=msg_id).execute()
        subject = next((header['value'] for header in message['payload']['headers'] if header['name'] == 'Subject'), "No Subject")
        print(f"Processing email with subject: {subject}")
        
        if term1.lower() in subject.lower():
            save_path = path1
            print(f"Subject contains '{term1}'. Saving to: {save_path}")
        elif term2.lower() in subject.lower():
            save_path = path2
            print(f"Subject contains '{term2}'. Saving to: {save_path}")
        else:
            print(f"Email subject doesn't contain '{term1}' or '{term2}'. Skipping.")
            return

        parts = message['payload'].get('parts', [])
        
        for part in parts:
            if part.get('filename'):
                filename = part['filename']
                print(f"Found attachment: {filename}")
                file_path = os.path.join(save_path, filename)
                
                # Check if file already exists
                if os.path.exists(file_path):
                    print(f"File {filename} already exists in {save_path}. Skipping download.")
                    continue

                if 'data' in part['body']:
                    data = part['body']['data']
                else:
                    print("Fetching attachment data...")
                    att_id = part['body']['attachmentId']
                    att = service.users().messages().attachments().get(userId=user_id, messageId=msg_id, id=att_id).execute()
                    data = att['data']
                
                file_data = base64.urlsafe_b64decode(data.encode('UTF-8'))
                
                print(f"Creating directory: {save_path}")
                os.makedirs(save_path, exist_ok=True)
                
                print(f"Saving attachment to: {file_path}")
                with open(file_path, 'wb') as f:
                    f.write(file_data)
                print(f"Successfully downloaded: {file_path}")

    except HttpError as error:
        print(f'An HTTP error occurred: {error}')
    except Exception as error:
        print(f'An unexpected error occurred: {error}')

def main():
    print("Starting Gmail Attachment Downloader")
    print("------------------------------------")

    service = get_gmail_service()
    sender_email = ['shyam.choudhary@sbcglobal.net']  # Replace with the sender's email address
    
    term1 = "EL CAMINO"  # Replace with actual term
    term2 = "MODESTO"  # Replace with actual term
    path1 = f"./{term1}"  # Replace with actual path
    path2 = f"./{term2}"  # Replace with actual path

    # Get start date from user
    start_date_str = input("Enter start date (MM/DD/YYYY): ")
    start_date = datetime.strptime(start_date_str, "%m/%d/%Y").date()

    # Get end date from user, use current date if blank
    end_date_str = input("Enter end date (MM/DD/YYYY) or press Enter for current date: ")
    if end_date_str:
        end_date = datetime.strptime(end_date_str, "%m/%d/%Y").date()
    else:
        end_date = date.today()
    
    print(f"Sender email: {sender_email}")
    print(f"Path 1: {path1}")
    print(f"Path 2: {path2}")
    print(f"Term 1: {term1}")
    print(f"Term 2: {term2}")
    print(f"Date range: {start_date.strftime('%m/%d/%Y')} to {end_date.strftime('%m/%d/%Y')}")
    print("------------------------------------")

    print("Creating main folders if they don't exist...")
    os.makedirs(path1, exist_ok=True)
    os.makedirs(path2, exist_ok=True)
    
    for i in sender_email:
        print(f"Searching for emails from: {i}")
        query = f'from:{i} after:{start_date.strftime("%Y/%m/%d")} before:{end_date.strftime("%Y/%m/%d")}'
        try:
            results = service.users().messages().list(userId='me', q=query).execute()
            messages = results.get('messages', [])

            if not messages:
                print('No messages found.')
            else:
                print(f'Number of messages found: {len(messages)}')
                for index, message in enumerate(messages, 1):
                    print(f"Processing message {index} of {len(messages)}")
                    get_attachments(service, 'me', message['id'], path1, path2, term1, term2)
        except HttpError as error:
            print(f'An HTTP error occurred while fetching messages: {error}')
        except Exception as error:
            print(f'An unexpected error occurred while fetching messages: {error}')

    print("Gmail Attachment Downloader completed.")
    
    # Process PDFs and update Excel
    print("\nStarting PDF processing and Excel update")
    print("------------------------------------")
    excel_path = "./FUEL DATA FORMAT.xlsx"
    process_and_update(path1, path2, excel_path)
    
    print("Program completed.")

if __name__ == '__main__':
    main()
    input("Press Enter to exit...")