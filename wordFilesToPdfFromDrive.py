import os
import io
import logging
import subprocess
import time
from datetime import datetime
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload
from google.auth.transport.requests import Request


def setup_logging():
    logging.basicConfig(
        filename='/Users/orkravitz/logs/word_to_pdf_conversion.log',
        level=logging.INFO,
        format='%(asctime)s:%(levelname)s:%(message)s'
    )
    logging.getLogger('googleapiclient.discovery_cache').setLevel(logging.ERROR)


def get_credentials(credentials_path, token_path, scopes):
    creds = None

    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(token_path, scopes)
        logging.info("Loaded existing token.")
    else:
        logging.info("No existing token found, need to authenticate.")

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
            logging.info("Refreshed the existing token.")
        else:
            flow = InstalledAppFlow.from_client_secrets_file(credentials_path, scopes)
            creds = flow.run_local_server(port=0)
            logging.info("Generated new token.")
        with open(token_path, 'w') as token_file:
            token_file.write(creds.to_json())
            logging.info("Saved new token.")

    return creds


def authenticate_google_drive(credentials_path, token_path, scopes):
    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(token_path, scopes)
        logging.info("Loaded existing token.")
    else:
        logging.info("No existing token found, need to authenticate.")
        flow = InstalledAppFlow.from_client_secrets_file(credentials_path, scopes)
        creds = flow.run_local_server(port=0)
        logging.info("Generated new token.")
        with open(token_path, 'w') as token_file:
            token_file.write(creds.to_json())
            logging.info("Saved new token.")
    return creds


def create_date_folder(service, parent_id):
    date_folder_name = datetime.today().strftime('%Y-%m-%d')
    query = f"name='{date_folder_name}' and mimeType='application/vnd.google-apps.folder' and '{parent_id}' in parents and trashed=false"
    res = service.files().list(q=query).execute()
    folders = res.get('files', [])
    if folders:
        return folders[0]['id']
    else:
        file_metadata = {'name': date_folder_name, 'mimeType': 'application/vnd.google-apps.folder', 'parents': [parent_id]}
        folder = service.files().create(body=file_metadata, fields='id').execute()
        return folder['id']


def get_file_owners(service, folder_id):
    # This function gets the owners of the files in the folder
    query = f"parents = '{folder_id}'"
    results = service.files().list(q=query, fields="files(id, name, owners)").execute()
    files = results.get('files', [])
    owners = {}
    for file in files:
        for owner in file.get('owners', []):
            owner_name = owner['displayName']
            owners[owner_name] = owners.get(owner_name, 0) + 1
    return owners


def convert_docx_to_pdf(docx_path, pdf_path):
    command = ['/Applications/LibreOffice.app/Contents/MacOS/soffice', '--headless', '--convert-to', 'pdf:writer_pdf_Export', '--outdir', os.path.dirname(pdf_path), docx_path]
    result = subprocess.run(command, capture_output=True, text=True)
    if result.returncode == 0:
        logging.info(f"Successfully converted {os.path.basename(docx_path)} to PDF at {pdf_path}")
        return True
    else:
        logging.error(f"Failed to convert {os.path.basename(docx_path)} to PDF: {result.stderr}")
        return False


def download_and_convert_files(service, source_folder_id, download_path, destination_path):
    results = service.files().list(q=f"'{source_folder_id}' in parents and trashed=false").execute()
    items = results.get('files', [])
    for item in items:
        file_id, file_name = item['id'], item['name']
        file_path = os.path.join(download_path, file_name)
        downloader = MediaIoBaseDownload(io.FileIO(file_path, 'wb'), service.files().get_media(fileId=file_id))
        done = False
        while not done:
            status, done = downloader.next_chunk()
        logging.info(f"Downloaded {file_name} to {file_path}")
        pdf_output = os.path.join(destination_path, file_name.replace('.docx', '.pdf'))
        if convert_docx_to_pdf(file_path, pdf_output):
            os.remove(file_path)
            try:
                service.files().delete(fileId=file_id).execute()
                logging.info(f"Successfully deleted {file_name}")
            except Exception as e:
                logging.error(f"Failed to delete {file_name}: {e}")


def upload_files(service, upload_path, folder_id):
    files_to_upload = [f for f in os.listdir(upload_path) if os.path.isfile(os.path.join(upload_path, f)) and not f.startswith('.') and f.lower().endswith('.pdf') and '.DS_Store' not in f and "subset" not in f]
    # Wait 30 seconds to ensure all background processes complete
    if len(files_to_upload) > 0:
        logging.info("Waiting for 30 seconds before uploading PDF files.")
        time.sleep(30)

        for pdf_file in files_to_upload:
            pdf_file_path = os.path.join(upload_path, pdf_file)
            file_metadata = {'name': pdf_file, 'parents': [folder_id]}
            media = MediaFileUpload(pdf_file_path, mimetype='application/pdf')
            file = service.files().create(body=file_metadata, media_body=media, fields='id').execute()
            os.remove(pdf_file_path)
            logging.info(f"Uploaded {pdf_file} to Google Drive and removed from local storage")


def check_folder_empty(serv, folder_id):
    res = serv.files().list(q=f"'{folder_id}' in parents and trashed=false").execute()
    files = res.get('files', [])
    return len(files) == 0


# Paths
SOURCE_ID = '18mPs5azAs4pSf-iE0IpR1sWFJb--fIs3'
DESTINATION_ID = '1CSAKqgTKfIrCOCZwiem8CQH-DRA6bqv1'
SOURCE_WORD_FILES = '/Users/orkravitz/Downloads/ProtectMyPDF/wordFilesToPdf'
DESTINATION_PDF_FILES = '/Users/orkravitz/Downloads/ProtectMyPDF/pdf_toSplit_toEncrypt'
SOURCE_UPLOAD_TO_GOOGLE_DRIVE = '/Users/orkravitz/Downloads/ProtectMyPDF/protected'

credentials_path = '/Users/orkravitz/Downloads/wordFilesToPdfCredentials/Credentials.json'
token_path = '/Users/orkravitz/Downloads/wordFilesToPdfCredentials/token.json'

SCOPES = ['https://www.googleapis.com/auth/drive']


def main():
    setup_logging()

    creds = get_credentials(credentials_path, token_path, SCOPES)

    service = build('drive', 'v3', credentials=creds)

    today_folder_id = create_date_folder(service, DESTINATION_ID)

    if check_folder_empty(service, SOURCE_ID):
        logging.info("No files to process in the Google Drive folder.")
    else:
        file_owners = get_file_owners(service, SOURCE_ID)
        for owner, count in file_owners.items():
            logging.info(f"Files uploaded by {owner}: {count}")
            download_and_convert_files(service, SOURCE_ID, SOURCE_WORD_FILES, DESTINATION_PDF_FILES)

    upload_files(service, SOURCE_UPLOAD_TO_GOOGLE_DRIVE, today_folder_id)


if __name__ == '__main__':
    main()