import os
from datetime import date
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseUpload
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

from config import CREDENTIALS_PATH, TOKEN_PATH


def get_drive_service():
    """Authenticates and returns a Google Drive service object."""
    scopes = ['https://www.googleapis.com/auth/drive']
    creds = None
    if os.path.exists(TOKEN_PATH):
        creds = Credentials.from_authorized_user_file(TOKEN_PATH, scopes)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_PATH, scopes)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_PATH, 'w') as token:
            token.write(creds.to_json())
    return build('drive', 'v3', credentials=creds)


def find_or_create_folder(service, name, parent_id):
    """Finds a folder by name inside parent_id, creates it if it doesn't exist."""
    query = f"'{parent_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false"
    results = service.files().list(q=query, fields="files(id, name)").execute()
    folders = results.get('files', [])
    for folder in folders:
        if name.lower() in folder['name'].lower():
            return folder['id']

    metadata = {
        'name': name,
        'mimeType': 'application/vnd.google-apps.folder',
        'parents': [parent_id]
    }
    folder = service.files().create(body=metadata, fields='id').execute()
    return folder['id']


def upload_to_drive(pdf_path, pn, vat_buffer=None):
    """
    Navigates Recycled Jewellery -> year -> Purchase Invoices -> month -> CPM -> PN folder
    (creating any missing folders along the way) and uploads the invoice PDF.
    Optionally uploads a VAT form from an in-memory BytesIO buffer.
    """
    service = get_drive_service()

    query = "name='Recycled Jewellery' and mimeType='application/vnd.google-apps.folder' and trashed=false"
    results = service.files().list(q=query, fields="files(id)").execute()
    folders = results.get('files', [])
    if not folders:
        raise Exception("Could not find 'Recycled Jewellery' folder in Drive")
    root_id = folders[0]['id']

    current_year = date.today().strftime('%Y')
    month_name = date.today().strftime('%B')
    pn_folder_name = f'PN{pn}'

    year_id = find_or_create_folder(service, current_year, root_id)
    purchase_id = find_or_create_folder(service, 'Purchase Invoices', year_id)
    month_id = find_or_create_folder(service, month_name, purchase_id)
    cpm_id = find_or_create_folder(service, 'CPM', month_id)
    pn_id = find_or_create_folder(service, pn_folder_name, cpm_id)

    file_metadata = {'name': f'{pn_folder_name}.pdf', 'parents': [pn_id]}
    media = MediaFileUpload(pdf_path, mimetype='application/pdf')
    service.files().create(body=file_metadata, media_body=media, fields='id').execute()
    print(f"Uploaded {pn_folder_name}.pdf to Drive successfully.")

    if vat_buffer is not None:
        vat_metadata = {'name': f'{pn_folder_name}_VAT264.pdf', 'parents': [pn_id]}
        media = MediaIoBaseUpload(vat_buffer, mimetype='application/pdf', resumable=False)
        service.files().create(body=vat_metadata, media_body=media, fields='id').execute()
        print(f"Uploaded VAT form for {pn_folder_name} to Drive successfully.")
