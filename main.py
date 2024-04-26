import os
import io
import logging
import subprocess
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from google.auth.transport.requests import Request

# Setup logging
logging.basicConfig(filename='/Users/orkravitz/logs/word_to_pdf_conversion.log', level=logging.INFO, format='%(asctime)s:%(levelname)s:%(message)s')

# Define the path to the credentials and token files
credentials_path = '/Users/orkravitz/Downloads/wordFilesToPdfCredentials/Credentials.json'
token_path = '/Users/orkravitz/Downloads/wordFilesToPdfCredentials/token.json'


def convert_docx_to_pdf(docx_path, pdf_path):
    try:
        # Build the pandoc command using list format
        command = ['pandoc', docx_path, '-o', pdf_path, '--pdf-engine=xelatex']
        # Run the command and capture output
        result = subprocess.run(command, capture_output=True, text=True)
        if result.returncode != 0:
            logging.error(f"Failed to convert {os.path.basename(docx_path)} to PDF: {result.stderr}")
            return False
        logging.info(f"Successfully converted {os.path.basename(docx_path)} to PDF.")
        return True
    except Exception as e:
        logging.error(f"Exception during conversion: {str(e)}")
        return False


# Authentication and service setup
SCOPES = ['https://www.googleapis.com/auth/drive.readonly']
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

service = build('drive', 'v3', credentials=creds)

# Download Word files from Google Drive
folder_id = '18mPs5azAs4pSf-iE0IpR1sWFJb--fIs3'
query = f"'{folder_id}' in parents"
results = service.files().list(q=query).execute()
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

    # Convert DOCX to PDF using the conversion function
    output_file = os.path.join(dest_path, file_name.replace('.docx', '.pdf'))
    if not convert_docx_to_pdf(file_path, output_file):
        logging.error(f"Failed to convert {file_name} to PDF.")

# Optional: Clean up source files
cleanup_result = os.system(f'rm -r {source_path}/*')
if cleanup_result == 0:
    logging.info(f"Cleaned up source files at {source_path}")
else:
    logging.error("Failed to clean up source files.")

logging.info("Script completed.")
