import os
import io
import logging
import subprocess
import time
import warnings

from datetime import datetime
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload
from google.auth.transport.requests import Request

warnings.filterwarnings('ignore', 'file_cache is only supported with oauth2client<4.0.0')
# Setup logging
logging.basicConfig(
    filename='/Users/orkravitz/logs/word_to_pdf_conversion.log',
    level=logging.INFO,
    format='%(asctime)s:%(levelname)s:%(message)s'
)
# Define the path to the credentials and token files
credentials_path = '/Users/orkravitz/Downloads/wordFilesToPdfCredentials/Credentials.json'
token_path = '/Users/orkravitz/Downloads/wordFilesToPdfCredentials/token.json'


def create_date_folder(serv, parent_id):
    date_folder_name = datetime.today().strftime('%Y-%m-%d')
    # Check if folder already exists
    query = f"name='{date_folder_name}' and mimeType='application/vnd.google-apps.folder' and '{parent_id}' in parents and trashed=false"
    res = serv.files().list(q=query).execute()
    folders = res.get('files', [])

    # Use the existing folder
    if folders:
        return folders[0]['id']
    else:
        # Create a new folder
        file_metadata = {'name': date_folder_name, 'mimeType': 'application/vnd.google-apps.folder',
                         'parents': [parent_id]}
        try:
            folder = serv.files().create(body=file_metadata, fields='id').execute()
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


def upload_file_to_drive(serv, filename, filepath, folder_id):
    if filename.startswith('.') or filename == '.DS_Store' or "subset" in filename:
        return False
    file_metadata = {'name': filename, 'parents': [folder_id]}
    media = MediaFileUpload(filepath, mimetype='application/pdf')
    try:
        file = serv.files().create(body=file_metadata, media_body=media, fields='id').execute()
        os.remove(filepath)
        logging.info(f"Uploaded {filename} to Google Drive and removed from local storage")
        return True
    except Exception as e:
        logging.error(f"Failed to upload {filename} to Google Drive: {str(e)}")
        return False


def check_folder_empty(serv, folder_id):
    res = serv.files().list(q=f"'{folder_id}' in parents and trashed=false").execute()
    files = res.get('files', [])
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

# Google Drive Path:
source_folder_id = '18mPs5azAs4pSf-iE0IpR1sWFJb--fIs3'
destination_folder_id = '1CSAKqgTKfIrCOCZwiem8CQH-DRA6bqv1'
# local encrypted file path:
protected_pdf_to_upload_path = "/Users/orkravitz/Downloads/ProtectMyPDF/protected"

# Initialize the Drive service for Google Drive operations
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
            try:
                service.files().delete(fileId=file_id).execute()
                logging.info(f"Successfully deleted {file_name}")
            except Exception as e:
                logging.error(f"Failed to delete {file_name}: {e}")
                # If possible, log e.response for more details
                if hasattr(e, 'response') and e.response:
                    logging.error(f"Detailed error response: {e.response}")

    # Wait 30 seconds to ensure all background processes complete
    logging.info("Waiting for 30 seconds before uploading PDF files.")
    time.sleep(30)


# List out all files in the directory
files_to_upload = [f for f in os.listdir(protected_pdf_to_upload_path)
                   if os.path.isfile(os.path.join(protected_pdf_to_upload_path, f)) and
                   not f.startswith('.') and f.lower().endswith('.pdf')]

# Check if the list is not empty, then proceed with upload
if files_to_upload:
    logging.info(f"Found {len(files_to_upload)} PDF file(s) to upload.")
    for pdf_file in files_to_upload:
        pdf_file_path = os.path.join(protected_pdf_to_upload_path, pdf_file)
        upload_file_to_drive(service, pdf_file, pdf_file_path, today_folder_id if today_folder_id else destination_folder_id)
else:
    logging.info("No PDF files found to upload.")

logging.info("Script completed.")