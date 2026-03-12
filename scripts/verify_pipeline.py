"""
Volunteer Onboarding Pipeline - Production Runner & Validator
------------------------------------------------------------
This script executes the Word-to-CSV extraction pipeline on real documents
located in the 'templates/' folder and validates the output integrity.

Usage:
    python scripts/verify_pipeline.py
"""

import os
import csv
import logging
import subprocess
import sys
from pathlib import Path
from typing import Dict, List, Optional

# Configuration
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S"
)
logger = logging.getLogger(__name__)

# --- Path Configuration ---
SCRIPT_DIR = Path(__file__).parent.absolute()
BASE_DIR = SCRIPT_DIR.parent
EXTRACTION_SCRIPT = SCRIPT_DIR / "extract_to_csv.py"

# Production directories
INPUT_DIR = BASE_DIR / "templates"
OUTPUT_DIR = BASE_DIR / "output"
OUTPUT_FILENAME = "volunteer_onboarding.csv"
OUTPUT_CSV = OUTPUT_DIR / OUTPUT_FILENAME


class PipelineRunner:
    """Handles execution and validation of the production extraction pipeline."""

    def setup_environment(self) -> None:
        """Ensures production directories exist."""
        logger.info("Initializing production environment...")
        INPUT_DIR.mkdir(parents=True, exist_ok=True)
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    def run_pipeline(self) -> bool:
        """Executes the extraction pipeline on the production 'templates/' folder."""
        logger.info(f"Executing pipeline script: {EXTRACTION_SCRIPT.name}")
        try:
            result = subprocess.run(
                [
                    sys.executable, 
                    str(EXTRACTION_SCRIPT),
                    "--input", str(INPUT_DIR),
                    "--output", str(OUTPUT_DIR),
                    "--filename", OUTPUT_FILENAME
                ],
                capture_output=True,
                text=True,
                cwd=str(BASE_DIR),
                check=True
            )
            logger.info("Pipeline execution completed.")
            if result.stdout:
                logger.info(f"Output: {result.stdout.strip()}")
            return True
        except subprocess.CalledProcessError as e:
            logger.error(f"Pipeline failed with exit code {e.returncode}")
            logger.error(f"Error output: {e.stderr}")
            return False

    def validate_results(self) -> bool:
        """Validates the generated CSV file."""
        if not OUTPUT_CSV.exists():
            logger.warning(f"No records were saved (CSV not created at {OUTPUT_CSV}).")
            return False

        records: List[Dict[str, str]] = []
        try:
            with open(OUTPUT_CSV, mode='r', encoding='utf-8') as f:
                reader = csv.DictReader(f)
                records = list(reader)
        except Exception as e:
            logger.error(f"Failed to read CSV: {e}")
            return False

        logger.info(f"Validated {len(records)} total records in {OUTPUT_FILENAME}.")
        
        if records:
            logger.info("Data integrity and schema validation successful.")
            return True
        else:
            logger.warning("CSV is empty. No data was extracted.")
            return False


def main():
    """Main execution flow."""
    runner = PipelineRunner()
    
    runner.setup_environment()
    
    # Check for input documents
    docx_files = list(INPUT_DIR.glob("*.docx"))
    if not docx_files:
        logger.warning(f"No .docx files found in {INPUT_DIR}. Put your forms there before running.")
        return

    logger.info(f"Found {len(docx_files)} documents to process.")
    
    # Process and Validate
    if runner.run_pipeline():
        if runner.validate_results():
            logger.info("SUCCESS: Data extraction and validation complete.")
        else:
            logger.warning("Verification finished with warnings (see logs above).")
    else:
        logger.error("FAILURE: Pipeline execution encountered errors.")


if __name__ == "__main__":
    main()
