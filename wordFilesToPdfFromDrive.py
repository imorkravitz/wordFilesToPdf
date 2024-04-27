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

# Setup logging
logging.basicConfig(filename='/Users/orkravitz/logs/word_to_pdf_conversion.log', level=logging.INFO, format='%(asctime)s:%(levelname)s:%(message)s')

# Define the path to the credentials and token files
credentials_path = '/Users/orkravitz/Downloads/wordFilesToPdfCredentials/Credentials.json'
token_path = '/Users/orkravitz/Downloads/wordFilesToPdfCredentials/token.json'


# Function to create a folder on Google Drive
# Function to create or find a folder with today's date
def create_date_folder(service, parent_id):
    date_folder_name = datetime.today().strftime('%Y-%m-%d')
    # Check if folder already exists
    query = f"name='{date_folder_name}' and mimeType='application/vnd.google-apps.folder' and '{parent_id}' in parents and trashed=false"
    results = service.files().list(q=query).execute()
    folders = results.get('files', [])

    if folders:
        return folders[0]['id']  # Use the existing folder
    else:
        # Create a new folder
        file_metadata = {'name': date_folder_name, 'mimeType': 'application/vnd.google-apps.folder',
                         'parents': [parent_id]}
        try:
            folder = service.files().create(body=file_metadata, fields='id').execute()
            return folder['id']
        except Exception as e:
            logging.error(f"Failed to create folder {date_folder_name}: {str(e)}")
            return None


def convert_docx_to_pdf(docx_path, pdf_path):
    try:
        command = ['/Applications/LibreOffice.app/Contents/MacOS/soffice', '--headless', '--convert-to', 'pdf:writer_pdf_Export', '--outdir', os.path.dirname(pdf_path), docx_path]
        result = subprocess.run(command, capture_output=True, text=True)
        if result.returncode == 0:
            logging.info(f"Successfully converted {os.path.basename(docx_path)} to PDF at {pdf_path}")
            return True
        else:
            logging.error(f"Failed to convert {os.path.basename(docx_path)} to PDF: {result.stderr}")
            return False
    except Exception as e:
        logging.error(f"Exception during conversion of {os.path.basename(docx_path)}: {str(e)}")
        return False


def upload_file_to_drive(service, filename, filepath, folder_id):
    if filename.startswith('.') or filename == '.DS_Store':  # Ignore system files
        return False
    file_metadata = {'name': filename, 'parents': [folder_id]}
    media = MediaFileUpload(filepath, mimetype='application/pdf')
    try:
        file = service.files().create(body=file_metadata, media_body=media, fields='id').execute()
        os.remove(filepath)
        logging.info(f"Uploaded {filename} to Google Drive and removed from local storage with ID {file['id']}")
        return True
    except Exception as e:
        logging.error(f"Failed to upload {filename} to Google Drive: {str(e)}")
        return False


def check_folder_empty(service, folder_id):
    results = service.files().list(q=f"'{folder_id}' in parents and trashed=false").execute()
    files = results.get('files', [])
    return len(files) == 0


# Authentication and service setup
SCOPES = ['https://www.googleapis.com/auth/drive']
creds = None
if os.path.exists(token_path):
    creds = Credentials.from_authorized_user_file(token_path, SCOPES)
    logging.info("Loaded existing token.")
else:
    logging.info("No existing token found, need to authenticate.")

if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
        logging.info("Refreshed the existing token.")
    else:
        flow = InstalledAppFlow.from_client_secrets_file(credentials_path, SCOPES)
        creds = flow.run_local_server(port=0)
        logging.info("Generated new token.")
    with open(token_path, 'w') as token_file:
        token_file.write(creds.to_json())
        logging.info("Saved new token.")


source_folder_id = '18mPs5azAs4pSf-iE0IpR1sWFJb--fIs3'
destination_folder_id = '1CSAKqgTKfIrCOCZwiem8CQH-DRA6bqv1'

service = build('drive', 'v3', credentials=creds)
today_folder_id = create_date_folder(service, destination_folder_id)  # Create or find today's folder


# Process files from source Google Drive folder
if check_folder_empty(service, source_folder_id):
    logging.info("No files to process in the Google Drive folder.")
else:
    results = service.files().list(q=f"'{source_folder_id}' in parents and trashed=false").execute()
    items = results.get('files', [])
    source_path = '/Users/orkravitz/Downloads/ProtectMyPDF/wordFilesToPdf'
    dest_path = '/Users/orkravitz/Downloads/ProtectMyPDF/pdf_toSplit_toEncrypt'
    protected_pdf_to_upload_path = "/Users/orkravitz/Downloads/ProtectMyPDF/protected"

    if not os.path.exists(source_path):
        os.makedirs(source_path)
        logging.info(f"Created directory for downloading Word files: {source_path}")

    for item in items:
        file_id = item['id']
        file_name = item['name']
        request = service.files().get_media(fileId=file_id)
        file_path = os.path.join(source_path, file_name)
        with io.FileIO(file_path, 'wb') as fh:
            downloader = MediaIoBaseDownload(fh, request)
            done = False
            while not done:
                status, done = downloader.next_chunk()
            logging.info(f"Downloaded {file_name} to {file_path}")

        output_file = os.path.join(dest_path, file_name.replace('.docx', '.pdf'))
        if convert_docx_to_pdf(file_path, output_file):
            os.remove(file_path)
            service.files().delete(fileId=file_id).execute()
            logging.info(f"Deleted {file_name} from Google Drive after conversion")

    # Wait 1 minute to ensure all background processes complete
    logging.info("Waiting for 1 minute before uploading PDF files.")
    time.sleep(60)

    # Upload all eligible PDFs from the protected path to today's folder in one bulk
    for pdf_file in os.listdir(protected_pdf_to_upload_path):
        pdf_file_path = os.path.join(protected_pdf_to_upload_path, pdf_file)
        if os.path.isfile(pdf_file_path) and not pdf_file.startswith('.'):
            upload_file_to_drive(service, pdf_file, pdf_file_path, today_folder_id if today_folder_id else destination_folder_id)

logging.info("Script completed.")