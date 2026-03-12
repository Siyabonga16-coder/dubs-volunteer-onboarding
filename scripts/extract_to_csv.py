import os
import csv
import argparse
import logging
from docx import Document
from pathlib import Path
from typing import Dict, List, Optional

# --- Configuration & Defaults ---
DEFAULT_INPUT_DIR = Path("templates")
DEFAULT_OUTPUT_DIR = Path("output")
DEFAULT_OUTPUT_FILENAME = "volunteer_onboarding.csv"

# Mapping of labels found in Word docs to our CSV fields
# Key: The label to search for (case-insensitive)
# Value: The corresponding column name in the CSV
FIELD_MAP = {
    "name": "Name",
    "surname": "Surname",
    "date of birth": "DateOfBirth",
    "cellphone": "Cellphone",
    "email": "Email"
}

# Setup Logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S"
)
logger = logging.getLogger(__name__)

def extract_text_from_docx(docx_path: Path) -> List[str]:
    """Extracts all text lines from paragraphs and tables in a docx file."""
    lines = []
    try:
        doc = Document(docx_path)
        # Extract from paragraphs
        for para in doc.paragraphs:
            if para.text.strip():
                lines.append(para.text.strip())
        
        # Extract from tables (common in forms)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    if cell.text.strip():
                        # Sometimes key/value are in the same cell or adjacent
                        lines.append(cell.text.strip())
    except Exception as e:
        logger.error(f"Failed to read {docx_path.name}: {e}")
    return lines

def parse_volunteer_data(lines: List[str], filename: str) -> Dict[str, str]:
    """Parses raw text lines into a structured volunteer data dictionary."""
    data = {field: "" for field in FIELD_MAP.values()}
    data["SourceFile"] = filename

    for line in lines:
        # Check each line against our field map
        for label, field in FIELD_MAP.items():
            # Match case-insensitively (e.g., "Name:", "NAME :", etc.)
            lower_line = line.lower()
            if label in lower_line:
                # Extract value after the first colon or space after the label
                parts = line.split(":", 1)
                if len(parts) > 1:
                    value = parts[1].strip()
                    # Only update if we found a non-empty value
                    if value and not data[field]:
                        data[field] = value
                else:
                    # Handle cases where there might not be a colon
                    # (e.g., field label and value in the same line/cell)
                    value = line[len(label):].strip().lstrip(":").strip()
                    if value and not data[field]:
                        data[field] = value
    return data

def save_to_csv(data_list: List[Dict[str, str]], output_path: Path):
    """Saves records to CSV, creating the file and header if needed."""
    if not data_list:
        return

    fieldnames = list(FIELD_MAP.values()) + ["SourceFile"]
    file_exists = output_path.exists()
    
    try:
        # 'a' mode to append new volunteers to the existing database
        with open(output_path, mode='a', newline='', encoding='utf-8') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            
            if not file_exists:
                writer.writeheader()
                logger.info(f"Created new CSV database: {output_path.name}")
            
            for data in data_list:
                writer.writerow(data)
        logger.info(f"Successfully saved {len(data_list)} records to {output_path.name}")
    except PermissionError:
        logger.error(f"Could not write to {output_path.name}. Please ensure the file is closed.")
    except Exception as e:
        logger.error(f"Error saving to CSV: {e}")

def main():
    parser = argparse.ArgumentParser(description="Professional volunteer data extraction utility.")
    parser.add_argument("--input", default=str(DEFAULT_INPUT_DIR), help="Directory with .docx forms")
    parser.add_argument("--output", default=str(DEFAULT_OUTPUT_DIR), help="Directory for CSV output")
    parser.add_argument("--filename", default=DEFAULT_OUTPUT_FILENAME, help="CSV filename")
    
    args = parser.parse_args()
    input_dir = Path(args.input)
    output_dir = Path(args.output)
    output_file = output_dir / args.filename

    # Ensure environment is ready
    input_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)

    logger.info(f"Scanning {input_dir.absolute()}...")
    docx_files = list(input_dir.glob("*.docx"))
    
    if not docx_files:
        logger.warning(f"No .docx files found in {input_dir}. Please add forms to process.")
        return

    extracted_data = []
    for docx_file in docx_files:
        logger.info(f"Processing: {docx_file.name}")
        text_lines = extract_text_from_docx(docx_file)
        data = parse_volunteer_data(text_lines, docx_file.name)
        
        # Only add if at least some data was found (prevents empty rows)
        if any(v for k, v in data.items() if k != "SourceFile"):
            extracted_data.append(data)
        else:
            logger.warning(f"No data found in {docx_file.name}. Ensure it matches the expected labels.")

    if extracted_data:
        save_to_csv(extracted_data, output_file)
    else:
        logger.info("No new data to extract.")

if __name__ == "__main__":
    main()

