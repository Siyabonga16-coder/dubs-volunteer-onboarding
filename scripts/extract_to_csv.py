"""
Volunteer Data Extraction Pipeline.
Extracts volunteer onboarding data from Word documents (.docx) into a CSV database.
"""

import os
import csv
import argparse
import logging
from docx import Document
from pathlib import Path
from typing import Dict, List, Optional

# Input and output directories
DEFAULT_INPUT_DIR = Path("templates")
DEFAULT_OUTPUT_DIR = Path("output")
DEFAULT_OUTPUT_FILENAME = "volunteer_onboarding.csv"

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

def get_existing_did_numbers(output_path: Path) -> set:
    """Read the existing CSV and return a set of all extracted DID Numbers."""
    existing_dids = set()
    if not output_path.exists():
        return existing_dids

    try:
        with open(output_path, mode='r', encoding='utf-8') as csvfile:
            reader = csv.DictReader(csvfile)
            for row in reader:
                did = row.get("DID Number", "").strip()
                if did:
                    existing_dids.add(did)
    except Exception as e:
        logger.error(f"Error reading existing DID numbers: {e}")
        
    return existing_dids

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
            
            # Write headers only if this is a brand new file
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
    # Setup CLI argument parsing
    parser = argparse.ArgumentParser(description="Volunteer Data Extraction Utility")
    parser.add_argument("--input", default=str(DEFAULT_INPUT_DIR), help="Input directory containing .docx forms")
    parser.add_argument("--output", default=str(DEFAULT_OUTPUT_DIR), help="Output directory for the CSV file")
    parser.add_argument("--filename", default=DEFAULT_OUTPUT_FILENAME, help="Output CSV filename")
    
    args = parser.parse_args()
    input_dir = Path(args.input)
    output_dir = Path(args.output)
    output_file = output_dir / args.filename

    # Ensure necessary input/output directories exist
    input_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)

    # Load existing records to prevent duplicates
    existing_dids = get_existing_did_numbers(output_file)
    if existing_dids:
        logger.info(f"Loaded {len(existing_dids)} existing records for deduplication.")

    logger.info(f"Scanning {input_dir.absolute()}...")
    
    # Locate all Word documents
    docx_files = list(input_dir.glob("*.docx"))
    if not docx_files:
        logger.warning(f"No .docx files found in {input_dir}. Please add forms to process.")
        return

    extracted_data = []
    
    # Iterate through all located documents
    for docx_file in docx_files:
        logger.info(f"Processing: {docx_file.name}")
        
        # Extract text and convert into structured data
        text_lines = extract_text_from_docx(docx_file)
        data = parse_volunteer_data(text_lines, docx_file.name)
        
        # Check for duplicates based on DID Number
        did_number = data.get("DID Number", "").strip()
        if did_number in existing_dids:
            logger.warning(f"Skipping duplicate: {docx_file.name} (DID: {did_number})")
            continue

        # Keep the data only if at least one target field was found (prevents empty rows)
        if any(data.values()):
            extracted_data.append(data)
            # Add to local set to prevent duplicates within the same run
            if did_number:
                existing_dids.add(did_number)
        else:
            logger.warning(f"No data found in {docx_file.name}. Ensure it matches the expected labels.")

    # Save aggregated data down to CSV
    if extracted_data:
        save_to_csv(extracted_data, output_file)
    else:
        logger.info("No new data to extract.")

if __name__ == "__main__":
    main()
