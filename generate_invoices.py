import os
import json
import io
import base64
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/gmail.send",
]
LOCAL_EXPORT_PATH = "/Users/aaron/Documents/217 Invoices"


def get_services():
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    if not creds or not creds.valid:
        flow = InstalledAppFlow.from_client_secrets_file("client_secret.json", SCOPES)
        creds = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(creds.to_json())
    return (
        build("drive", "v3", credentials=creds),
        build("sheets", "v4", credentials=creds),
        build("gmail", "v1", credentials=creds),
    )


def find_folder(service, athlete_name):
    query = f"name = '{athlete_name}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
    results = service.files().list(q=query, fields="files(id, name)").execute()
    files = results.get("files", [])
    return files[0]["id"] if files else None


def find_template_invoice(service, folder_id):
    query = f"'{folder_id}' in parents and name contains 'IN-' and trashed = false"
    results = service.files().list(q=query, fields="files(id, name)").execute()
    files = results.get("files", [])
    return files[0]["id"] if files else None


def send_invoice_email(service, to_email, subject, file_path):
    """Creates and sends an email with proper name and PDF attachment."""
    message = MIMEMultipart()
    message["to"] = to_email
    # 1. FIX: Format the 'From' header to show your name
    message["from"] = "Aaron Manning <aaron.manning613@gmail.com>"
    message["subject"] = subject

    # 2. FIX: Add a tiny bit of body text so the PDF stays as an attachment
    # Even a blank string or a "Please find attached" helps the email client
    body = "Please find your coaching invoice attached."
    message.attach(MIMEText(body, "plain"))

    # Attachment logic
    with open(file_path, "rb") as f:
        part = MIMEApplication(f.read(), _subtype="pdf")
        # Ensure the filename is set correctly for the UI
        part.add_header(
            "Content-Disposition", "attachment", filename=os.path.basename(file_path)
        )
        message.attach(part)

    # Encode and send
    raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode()
    service.users().messages().send(userId="me", body={"raw": raw_message}).execute()


def main():
    drive_service, sheets_service, gmail_service = get_services()

    target_month_str = input("Enter the month name (e.g., March): ")
    current_inv_num = int(input("Enter the starting invoice number (e.g., 111): "))

    try:
        month_dt = datetime.strptime(target_month_str, "%B")
        invoice_date = f"{month_dt.month}/1/2026"
        month_abbr = target_month_str[:3]
    except ValueError:
        print("Error: Invalid month name.")
        return

    with open("athletes.json", "r") as f:
        athletes = json.load(f)

    if not os.path.exists(LOCAL_EXPORT_PATH):
        os.makedirs(LOCAL_EXPORT_PATH)

    for athlete in athletes:
        name = athlete["name"]
        email = athlete["email"]
        print(f"\n--- Processing {name} ---")

        folder_id = find_folder(drive_service, name)
        template_id = find_template_invoice(drive_service, folder_id)

        if folder_id and template_id:
            inv_string = f"IN-{current_inv_num:04d}"
            new_filename = f"{inv_string}_{month_abbr}26"

            # 1. Copy
            copy_metadata = {"name": new_filename, "parents": [folder_id]}
            new_file = (
                drive_service.files()
                .copy(fileId=template_id, body=copy_metadata)
                .execute()
            )
            new_file_id = new_file.get("id")

            # 2. Edit
            batch_update_body = {
                "valueInputOption": "USER_ENTERED",
                "data": [
                    {"range": "G6", "values": [[invoice_date]]},
                    {"range": "G8", "values": [[inv_string]]},
                ],
            }
            sheets_service.spreadsheets().values().batchUpdate(
                spreadsheetId=new_file_id, body=batch_update_body
            ).execute()

            # 3. Export PDF
            local_file_path = os.path.join(LOCAL_EXPORT_PATH, f"{new_filename}.pdf")
            request = drive_service.files().export_media(
                fileId=new_file_id, mimeType="application/pdf"
            )
            with io.FileIO(local_file_path, "wb") as fh:
                downloader = MediaIoBaseDownload(fh, request)
                done = False
                while done is False:
                    status, done = downloader.next_chunk()

            # 4. Email
            print(f"  [~] Sending email to {email}...")
            subject = f"{target_month_str} coaching invoice"
            send_invoice_email(gmail_service, email, subject, local_file_path)

            print(f"  [+] Success: Invoice {inv_string} sent.")

            current_inv_num += 1

    print("\nAll tasks complete. Invoices generated and emailed.")


if __name__ == "__main__":
    main()
