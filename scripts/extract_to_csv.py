"""
Volunteer Data Extraction Pipeline.
Extracts volunteer onboarding data from Word documents (.docx) into a CSV database.
"""
import io
import os
import csv
import argparse
import logging
#from docx import Document
from pathlib import Path
from typing import Dict, List, Optional
from urllib.request import Request

from docx import  Document
from pip._internal.network.auth import Credentials

from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload

# Input and output directories
DEFAULT_INPUT_DIR = Path("templates")
DEFAULT_OUTPUT_DIR = Path("output")
DEFAULT_OUTPUT_FILENAME = "volunteer_onboarding.csv"

Scopes = ["https://www.googleapis.com/auth/drive"]#asking for writing and readimg accesss uding google api

# Map of Word document labels to CSV column names.
# Made keys lowercase for case-insensitive matching during extraction.
FIELD_MAP = {
    "dubs impact driver (did) no.": "DID Number",
    "country": "Country",
    "province": "Province",
    "street address": "Street Address",
    "suburb": "Suburb/Area",
}

# Log format includes time, severity level, and standard messaging.
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S"
)
logger = logging.getLogger(__name__)


#LOGGING INTO Google drive which is the main process
def authenticate_drive():
    creds = None# Log in Google drive and return the servive object

    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', Scopes)

    if not creds or not creds.valid:# checking they are not available or invalid
        if creds and creds.expired and creds.refresh_token:# if maybe expired
            creds.refresh(Request())# refresh in background if they expired
        else:
            #if it the first time ,ask the user to login
            flow  = InstalledAppFlow.from_client_secrets_file('credentials.json', Scopes)
            creds = flow.run_local_server(port=0)

        #save login details avoiding loging in next time
        with open('token.json', 'w') as token:
            token .write(creds.to_json())
    # return object for connection to google drive
    return  build('drive','v3',credentials=creds)

 #downlod all word documents found on the googgle drive
def download_docs_from_drive(service, folder_id: str, download_dir: Path):
    download_dir.mkdir(parents=True, exist_ok=True)
    query = f"'{folder_id}'in parents and mimeType='application/vnd.openxmlformats-officedocument.wordprocessingml.document'"

    results = service.files().list(q=query,fields="files(id,name)").execute()

    items = results.get('files', [])# stracting lists from google files

    # if no lists extracted output the message
    if not items:
        logger.warning("no words documets found!!...")
        return

    for item in items:
        file_id = item['id']
        file_name=item['name'] # get actully name oof the file
        local_path = download_dir / file_name

        logger.info(f"Downloding from drive .......: {local_path}")
        request = service.file().get_media(fileId=file_id)#requesting for the exact content download
        fh = io.BytesIO()#create a mempty memory buffer to hold incoming data

        downloader = MediaIoBaseDownload(fh,request)#setting the downloader to pull data

        #Start looping in download the file ,and this prevents craishing in big files
        done = False
        while done is False:
            status,done = downloader.next_chuck()

        with  open(local_path, 'wb') as f:
            f.write(fh.getvalue())






def extract_text_from_docx(docx_path: Path) -> List[str]:
    """Extract text from all paragraphs and tables within a Word document."""
    lines = []
    try:
        doc = Document(docx_path)
        
        # 1. Read standard text paragraphs
        for para in doc.paragraphs:
            if para.text.strip():
                lines.append(para.text.strip())
        
        # 2. Read structured form tables
        for table in doc.tables:
            for row in table.rows:
                # Track unique cells per row to ignore merged cell overlaps
                unique_cells = []
                for cell in row.cells:
                    # Clean up random newline characters inside table cells
                    text = cell.text.strip().replace('\n', ' ')
                    if text and (not unique_cells or unique_cells[-1] != text):
                        unique_cells.append(text)
                
                # If a row has two cells, treat them as a 'Label: Value' pair
                if len(unique_cells) >= 2:
                    lines.append(f"{unique_cells[0]}: {unique_cells[1]}")
                # Otherwise, treat it as a single flat string
                elif len(unique_cells) == 1:
                    lines.append(unique_cells[0])
                    
    except Exception as e:
        logger.error(f"Failed to read {docx_path.name}: {e}")
        
    return lines

def parse_volunteer_data(lines: List[str], filename: str) -> Dict[str, str]:
    """Parse raw text lines into a structured dictionary matching our CSV schema."""
    # Initialize dictionary with empty defaults
    data = {field: "" for field in FIELD_MAP.values()}

    for line in lines:
        lower_line = line.lower()
        
        # Check if any mapped label exists in the current line
        for label, field in FIELD_MAP.items():
            if label in lower_line:
                # Attempt to split the line by a colon delimiter (e.g., 'Name: John')
                parts = line.split(":", 1)
                
                if len(parts) > 1:
                    value = parts[1].strip()
                    # Assign the parsed value if we don't already have one
                    if value and not data[field]:
                        data[field] = value
                else:
                    # Fallback: Strip the label text directly if no colon is found (e.g., 'Name John')
                    value = line[len(label):].strip().lstrip(":").strip()
                    if value and not data[field]:
                        data[field] = value
                        
    return data

def save_to_csv(data_list: List[Dict[str, str]], output_path: Path):
    """Append extracted volunteer records to the final CSV database."""
    # Skip processing if no data was found
    if not data_list:
        return

    fieldnames = list(FIELD_MAP.values())
    file_exists = output_path.exists()
    
    try:
        # Open the CSV in append mode ('a') to safely add new rows
        with open(output_path, mode='a', newline='', encoding='utf-8') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            
            #Write headers only if this is a brandnew file
            if not file_exists:
                writer.writeheader()
                logger.info(f"Created new CSV database: {output_path.name}")
            
            # Write each volunteer record as a row
            for data in data_list:
                writer.writerow(data)
                
        logger.info(f"Successfully saved {len(data_list)} records to {output_path.name}")
        
    except PermissionError:
        logger.error(f"Could not write to {output_path.name}. Ensure it is not open in another program.")
    except Exception as e:
        logger.error(f"Error saving to CSV: {e}")

def main():
    parser = argparse.ArgumentParser(description="Volunteer Data Extraction Utility")
    parser.add_argument("--output", default=str(DEFAULT_OUTPUT_DIR), help="Output directory for the CSV file")
    parser.add_argument("--filename", default=DEFAULT_OUTPUT_FILENAME, help="Output CSV filename")

    args = parser.parse_args()
    output_dir = Path(args.output)
    output_file = output_dir / args.filename

    output_dir.mkdir(parents=True, exist_ok=True)

    DRIVE_FOLDER_ID = '1ExifgMx_ivxirv3LnO4nS9_9HC9M4CVr'



    logger.info("Authenticating with Google Drive...")
    service = authenticate_drive()
    logger.info("Fetching files from Google Drive...")
    download_docs_from_drive(service,DRIVE_FOLDER_ID,DEFAULT_OUTPUT_DIR)
    logger.info(f"Scanning downloaded files in {DEFAULT_INPUT_DIR.absolute()}...")
    # Locate all Word documents
    docx_files = list(DEFAULT_INPUT_DIR.glob("*.docx"))
    if not docx_files:
        logger.warning(f"No .docx files found in Drive folder. ending program....")
        return

    extracted_data = []
    
    # Iterate through all located documents
    for docx_file in docx_files:
        logger.info(f"Processing: {docx_file.name}")
        
        # Extract text and convert into structured data
        text_lines = extract_text_from_docx(docx_file)
        data = parse_volunteer_data(text_lines, docx_file.name)
        
        # Keep the data only if at least one target field was found (prevents empty rows)
        if any(data.values()):
            extracted_data.append(data)
        else:
            logger.warning(f"No data found in {docx_file.name}. Ensure it matches the expected labels.")

    # Save aggregated data down to CSV
    if extracted_data:
        save_to_csv(extracted_data, output_file)
    else:
        logger.info("No new data to extract.")

if __name__ == "__main__":
    main()
          

