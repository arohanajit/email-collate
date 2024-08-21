import csv
import os
import base64
import mimetypes
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

# Gmail API Scopes
SCOPES = ['https://www.googleapis.com/auth/gmail.send']

def get_gmail_service():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    service = build('gmail', 'v1', credentials=creds)
    return service

def create_message_with_attachments(sender, to, subject, body, file_paths):
    message = MIMEMultipart()
    message['to'] = to
    message['from'] = sender
    message['subject'] = subject

    msg = MIMEText(body, 'html')
    message.attach(msg)

    for file_path in file_paths:
        content_type, encoding = mimetypes.guess_type(file_path)

        if content_type is None or encoding is not None:
            content_type = 'application/octet-stream'
        main_type, sub_type = content_type.split('/', 1)

        with open(file_path, 'rb') as file:
            attachment = MIMEBase(main_type, sub_type)
            attachment.set_payload(file.read())

        encoders.encode_base64(attachment)
        filename = os.path.basename(file_path)
        attachment.add_header('Content-Disposition', 'attachment', filename=filename)
        message.attach(attachment)

    raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode("utf-8")
    return {'raw': raw_message}

def send_email(service, message):
    try:
        message = service.users().messages().send(userId="me", body=message).execute()
        print(f'Email sent successfully. Message Id: {message["id"]}')
    except Exception as e:
        print(f'An error occurred: {str(e)}')

def main():
    service = get_gmail_service()
    sender_email = "aajit@ncsu.edu"  # Replace with your Gmail address
    
    # Paths to the attachments
    path1 = '/Users/arohanajit/Documents/resume/word/arohanajit-sde.pdf'  # Replace with the actual path to your first file
    path2 = '/Users/arohanajit/Documents/resume/word/arohanajit-ml.pdf'  # Replace with the actual path to your second file

    with open('uploaded_professors.csv', mode='r') as csv_file:
        csv_reader = csv.DictReader(csv_file)
        for row in csv_reader:
            name = row['name']
            email = row['email']
            subject = "Interest in Student Jobs"
            body = f"""
<html>
  <body>
    <p>Dear Professor {name.split()[-1]},</p>
    <p>I am Arohan Ajit, a Master of Science in Computer Science student at North Carolina State University (GPA: 3.8/4.0). I am writing to express my interest in any available positions as a coder, teaching assistant, research assistant, or grader in your department.</p>
    <p>Key qualifications:</p>
    <ul>
      <li>Software Development: 2 years experience at Accenture, current internship at Chirpn
        <ul>
          <li>At Chirpn: Optimized API endpoints, improved data processing efficiency, and implemented job scheduling mechanisms</li>
        </ul>
      </li>
      <li>Machine Learning: 2 years experience at Omdena, current ML internship at Chirpn
        <ul>
          <li>At Chirpn: Integrated advanced AI models, expanded vector database capabilities, and optimized document synchronization</li>
        </ul>
      </li>
      <li>Strong background in both full-stack development and ML model deployment</li>
      <li>Proficient in Python, ReactJS, Node.js, TensorFlow, PyTorch, and cloud platforms (AWS, Azure)</li>
    </ul>
    <p>I am eager to contribute my skills to your projects or courses. Please let me know if there are any opportunities available or if you need any additional information.</p>
    <p>Thank you for your consideration.</p>
    <p>Best regards,<br>Arohan Ajit</p>
  </body>
</html>"""
            message = create_message_with_attachments(sender_email, email, subject, body, [path1, path2])
            send_email(service, message)

if __name__ == '__main__':
    main()