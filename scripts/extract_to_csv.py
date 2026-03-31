"""
Volunteer Data Onboarding - Extraction Pipeline.

This script parses volunteer onboarding Microsoft Word documents (.docx) and 
converts the structured form data into a neat, searchable CSV database.
It handles highlighting, grid-based cell joining, and strict label matching.
"""

import os
import csv
import argparse
import logging
from docx import Document
from pathlib import Path
from typing import Dict, List, Optional, Set

# =========================================================================
# CONFIGURATION & CONSTANTS
# =========================================================================

# Default system paths
DEFAULT_INPUT_DIR = Path("C:\\Users\\hlong\\Music\\CSV\\dubs-volunteer-onboarding\\data\\input")
DEFAULT_OUTPUT_DIR = Path("C:\\Users\\hlong\\Music\\CSV\\dubs-volunteer-onboarding\\data\\output")
DEFAULT_OUTPUT_FILENAME = "volunteer_onboarding.csv"
DEFAULT_FULL_FILENAME = "volunteer_onboarding_full.csv"

# FIELD_MAP: Defines the essential columns for the primary "Filtered" CSV output.
FIELD_MAP = {
    "dubs impact driver (did) no.": "DID Number",
    "country": "Country",
    "province / state": "Province/State",
    "street address": "Street Address",
    "suburb": "Suburb/Area",
}

# FULL_FIELD_MAP: The master definition for all fields captured in the "Full" CSV.
# The keys represent the labels found in the .docx files (case-insensitive prefixes).
# The values represent the final CSV header names.
FULL_FIELD_MAP = {
    # 1. Personal Information
    "dubs impact driver (did) no.": "DID Number",
    "first name": "First Name",
    "last name": "Last Name",
    "second name": "Second Name",
    "gender": "Gender",
    "phone number": "Phone Number",
    "race": "Race",
    "citizenship": "Citizenship",
    "foreign citizenship": "Foreign Citizenship",
    "home language": "Home Language",
    "date of birth (yyyymmdd)": "Date of Birth",
    "pronouns (optional)": "Pronouns",
    "name pronunciation(optional)": "Name Pronunciation",
    
    # 2. Contact Information
    "primary cellphone number": "Primary Cellphone Number",
    "whatsapp cellphone number": "WhatsApp Cellphone Number",
    "alternate cellphone number": "Alternate Cellphone Number",
    "personal email address": "Personal Email Address",
    
    # 3. Location Information
    "country": "Country",
    "province / state": "Province/State",
    "city": "City",
    "street address": "Street Address",
    "suburb / area": "Suburb/Area",
    
    # 4. Next of Kin
    "next of kin full name": "Next of Kin Full Name",
    "next of kin cellphone number": "Next of Kin Cellphone Number",
    
    # 5. Education
    "highest qualification": "Highest Qualification",
    "institution": "Institution",
    "qualification name": "Qualification Name",
    "year obtained": "Year Obtained",
    "former secondary/high school name": "Former School Name",
    
    # 6. Certifications (Explicitly numbered for consistent ordering)
    "certification 001 - certification name": "Certification 001 - Certification Name",
    "certification 001 - issuing institution / provider": "Certification 001 - Issuing Institution / Provider",
    "certification 001 - year obtained (optional)": "Certification 001 - Year Obtained (optional)",
    "certification 002 - certification name": "Certification 002 - Certification Name",
    "certification 002 - issuing institution / provider": "Certification 002 - Issuing Institution / Provider",
    "certification 002 - year obtained (optional)": "Certification 002 - Year Obtained (optional)",
    "certification 003 - certification name": "Certification 003 - Certification Name",
    "certification 003 - issuing institution / provider": "Certification 003 - Issuing Institution / Provider",
    "certification 003 - year obtained (optional)": "Certification 003 - Year Obtained (optional)",
    "certification 004 - certification name": "Certification 004 - Certification Name",
    "certification 004 - issuing institution / provider": "Certification 004 - Issuing Institution / Provider",
    "certification 004 - year obtained (optional)": "Certification 004 - Year Obtained (optional)",
    "certification 005 - certification name": "Certification 005 - Certification Name",
    "certification 005 - issuing institution / provider": "Certification 005 - Issuing Institution / Provider",
    "certification 005 - year obtained (optional)": "Certification 005 - Year Obtained (optional)",
    
    # 7. Work / Role
    "team": "Team",
    "division": "Division",
    "role": "Role",
    "reporting officer/line manager": "Reporting Manager",
    "date joined (yyyymmdd)": "Date Joined",
    
    # 8. Skills & Interests
    "top technical skills": "Top Technical Skills",
    "top soft skills": "Top Soft Skills",
    "favourite quote or line": "Favourite Quote",
    "hobbies & interests": "Hobbies & Interests",
}

# Logger setup
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S"
)
logger = logging.getLogger(__name__)


# =========================================================================
# WORD DOCUMENT PARSING HELPERS
# =========================================================================

def is_cell_highlighted(cell) -> bool:
    """Detect if a Word table cell contains highlighted text markers."""
    for para in cell.paragraphs:
        for run in para.runs:
            if run.font.highlight_color is not None:
                return True
    return False


def _process_table_row(row, context: Dict[str, str]) -> List[str]:
    """
    Apply heuristics to extract meaningful Label: Value pairs from a table row.
    Heuristics include grid joining, highlight detection, and context labels (Certifications).
    """
    row_lines = []
    cells_info = []
    
    # Pre-scan cells for text and highlight status
    for cell in row.cells:
        text = cell.text.strip().replace('\n', ' ')
        # Deduplicate text from merged cells in the same row
        if text and (not cells_info or cells_info[-1]['text'] != text):
            cells_info.append({
                'text': text,
                'is_highlighted': is_cell_highlighted(cell)
            })
            
    if not cells_info:
        return row_lines

    cells_text = [c['text'] for c in cells_info]
    highlighted_cells = [c['text'] for c in cells_info if c['is_highlighted']]

    # Heuristic 1: Multiple choice selection (e.g. Citizenship, Gender, Race)
    # If a row has a known label and one option marked with a highlight, select that option.
    if len(cells_info) > 1:
        label_lower = cells_info[0]['text'].lower()
        if any(key in label_lower for key in ["citizenship", "gender", "race"]):
            if highlighted_cells:
                row_lines.append(f"{cells_info[0]['text']}: {highlighted_cells[0]}")
                return row_lines

    # Heuristic 2: Digit Grids (e.g. Date Joined, DID No.)
    # If a label is followed by many separate cells (digits), join them into a single string.
    if len(cells_text) >= 3:
        label_0 = cells_info[0]['text'].lower()
        if "date" in label_0 or "did" in label_0:
            value = "".join([c['text'] for c in cells_info[1:] if c['text'] != ' ']).strip()
            row_lines.append(f"{cells_info[0]['text']}: {value}")
            return row_lines
        
        # Standard Grid (Label | Value...) with context persistence
        context['label'] = cells_text[0]
        sub_label = cells_text[1]
        value = " ".join(cells_text[2:])
        # Prioritize highlighting even in multi-cell grids
        if highlighted_cells and highlighted_cells[0] != context['label']:
            value = highlighted_cells[0]
            
        row_lines.append(f"{context['label']} - {sub_label}: {value}")
    
    # Heuristic 3: Simple Label-Value Pairs (2 cells)
    elif len(cells_text) == 2:
        label, value = cells_text[0], cells_text[1]
        
        # Certification blocks often use "Certification Name" as a sub-label
        if context['label'] and label in ["Certification Name", "Issuing Institution / Provider"]:
            row_lines.append(f"{context['label']} - {label}: {value}")
        else:
            if label.startswith("Certification"):
                context['label'] = label
            row_lines.append(f"{label}: {value}")
            
    # Heuristic 4: Section Headings (1 cell)
    elif len(cells_text) == 1:
        label = cells_text[0]
        if label.startswith("Certification"):
            context['label'] = label
        row_lines.append(label)
        
    return row_lines


def extract_text_from_docx(docx_path: Path) -> List[str]:
    """
    Read a Word document and output a collection of 'Label: Value' strings.
    Orchestrates paragraph extraction and complex table heuristics.
    """
    lines = []
    try:
        doc = Document(docx_path)
        
        # 1. Capture standard paragraphs (usually instructions or simple lines)
        for para in doc.paragraphs:
            if para.text.strip():
                lines.append(para.text.strip())
        
        # 2. Capture table data (using stateful context for nested labels)
        context = {'label': ""}
        for table in doc.tables:
            for row in table.rows:
                lines.extend(_process_table_row(row, context))
                    
    except Exception as e:
        logger.error(f"Failed to read {docx_path.name}: {e}")
        
    return lines


# =========================================================================
# DATA STRUCTURING & CSV LOGIC
# =========================================================================

def parse_volunteer_data(lines: List[str], filename: str) -> Dict[str, str]:
    """
    Convert raw extracted lines into a structured data record (dictionary).
    Uses strict label matching and prefix support for long form labels.
    """
    # Initialize record with "null" defaults to maintain CSV structure
    data = {field: "null" for field in FULL_FIELD_MAP.values()}

    for line in lines:
        if ":" not in line:
            continue
            
        parts = line.split(":", 1)
        raw_label = parts[0].strip().lower()
        value = parts[1].strip()
        
        if not value:
            continue
            
        # Match the extracted label against our master field map
        # We use 'startswith' to handle long instructional labels (e.g. Skills section)
        for label, field in FULL_FIELD_MAP.items():
            if raw_label.startswith(label.lower()) or raw_label.startswith(field.lower()):
                # Only populate if not already set (prevents secondary matches from overwriting)
                if data[field] == "null":
                    data[field] = value
                break
        
    return data


def save_to_csv(data_list: List[Dict[str, str]], output_path: Path, fieldnames: List[str], pad_columns: bool = True):
    """
    Write processed records to a CSV file.
    'pad_columns' ensures the CSV remains human-readable in flat text editors (Neat Mode).
    """
    if not data_list:
        return

    # Calculate optimal column padding across all records
    column_widths = {}
    if pad_columns:
        for field in fieldnames:
            max_w = len(str(field))
            for data in data_list:
                val_w = len(str(data.get(field, "null")))
                max_w = max(max_w, val_w)
            column_widths[field] = max_w

    try:
        # Overwrite file to ensure padding consistency for ALL rows
        with open(output_path, mode='w', newline='', encoding='utf-8') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames, extrasaction='ignore', quoting=csv.QUOTE_MINIMAL)
            
            # Format and write the header row
            header_row = {f: (f.ljust(column_widths[f]) if pad_columns else f) for f in fieldnames}
            writer.writerow(header_row)
            
            # Format and write data rows
            for data in data_list:
                row = {}
                for field in fieldnames:
                    val = str(data.get(field, "null"))
                    row[field] = val.ljust(column_widths[field]) if pad_columns else val
                writer.writerow(row)
                
        logger.info(f"Saved {len(data_list)} records to {output_path.name} (Neat mode: {pad_columns})")
        
    except (PermissionError, Exception) as e:
        logger.error(f"Error saving to {output_path.name}: {e}")


def load_existing_records(output_path: Path) -> List[Dict[str, str]]:
    """Recover previously extracted records from an existing CSV files."""
    records = []
    if not output_path.exists():
        return records

    try:
        with open(output_path, mode='r', encoding='utf-8') as csvfile:
            reader = csv.DictReader(csvfile)
            for row in reader:
                # Clean up neat-mode padding for internal processing
                clean_row = {k.strip(): v.strip() for k, v in row.items()}
                records.append(clean_row)
    except Exception as e:
        logger.error(f"Error reading existing records: {e}")
        
    return records


# =========================================================================
# MAIN PIPELINE WORKFLOW
# =========================================================================

def process_onboarding_directory(input_dir: Path, output_file: Path, full_output_file: Path):
    """Orchestrates the scanning, parsing, deduplication, and saving of all forms."""
    
    # 1. Load existing data for deduplication (source of truth is the DID Number)
    all_filtered_records = load_existing_records(output_file)
    all_full_records = load_existing_records(full_output_file)
    existing_dids = {r.get("DID Number", "").strip() for r in all_filtered_records if r.get("DID Number")}
    
    if existing_dids:
        logger.info(f"Loaded {len(existing_dids)} records for deduplication check.")

    # 2. Identify new files to process
    docx_files = list(input_dir.glob("*.docx"))
    if not docx_files and not all_filtered_records:
        logger.warning(f"No .docx files found in {input_dir}.")
        return

    # 3. Batch process new documents
    new_count = 0
    for docx_file in docx_files:
        lines = extract_text_from_docx(docx_file)
        data = parse_volunteer_data(lines, docx_file.name)
        
        did = data.get("DID Number", "null").strip()
        if did != "null" and did in existing_dids:
            continue

        # Valid records must contain at least one piece of mapped volunteer data
        if any(v != "null" for k, v in data.items() if k in FULL_FIELD_MAP.values()):
            all_filtered_records.append(data)
            all_full_records.append(data)
            if did != "null":
                existing_dids.add(did)
            new_count += 1
            logger.info(f"Extracted: {docx_file.name}")

    # 4. Consolidate and save all results
    if all_filtered_records:
        # Save primary filtered view
        save_to_csv(all_filtered_records, output_file, list(FIELD_MAP.values()))
        
        # Save comprehensive full view
        full_headers = list(dict.fromkeys(FULL_FIELD_MAP.values()))
        save_to_csv(all_full_records, full_output_file, full_headers)
        
        summary = f"Done! {new_count} new records added." if new_count > 0 else "CSV files refreshed."
        logger.info(summary)


def main():
    """CLI entry point for the extraction utility."""
    parser = argparse.ArgumentParser(description="Volunteer Data Extraction Utility")
    parser.add_argument("--input", default=str(DEFAULT_INPUT_DIR), help="Source directory for .docx forms")
    parser.add_argument("--output", default=str(DEFAULT_OUTPUT_DIR), help="Destination directory for CSVs")
    parser.add_argument("--filename", default=DEFAULT_OUTPUT_FILENAME, help="Filtered CSV filename")
    parser.add_argument("--full-filename", default=DEFAULT_FULL_FILENAME, help="Full dataset filename")
    
    args = parser.parse_args()
    process_onboarding_directory(Path(args.input), Path(args.output) / args.filename, Path(args.output) / args.full_filename)


if __name__ == "__main__":
    main()
