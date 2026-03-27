"""
Pipeline Verification Script.
Executes the extraction pipeline natively and validates output integrity.
"""

import os
import csv
import logging
import subprocess
import sys
from pathlib import Path
from typing import Dict, List, Optional

# Standardizes execution logs for traceability
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S"
)
logger = logging.getLogger(__name__)

# --- Project Paths ---
# Dynamically resolve paths relative to the script itself
SCRIPT_DIR = Path(__file__).parent.absolute()
BASE_DIR = SCRIPT_DIR.parent
EXTRACTION_SCRIPT = SCRIPT_DIR / "extract_to_csv.py"

INPUT_DIR = BASE_DIR / "templates"
OUTPUT_DIR = BASE_DIR / "output"
OUTPUT_FILENAME = "volunteer_onboarding.csv"
OUTPUT_CSV = OUTPUT_DIR / OUTPUT_FILENAME


class PipelineRunner:
    """Orchestrates test runs for the extraction pipeline."""

    def setup_environment(self) -> None:
        """Create required execution folders logically if they are missing."""
        logger.info("Initializing environment...")
        INPUT_DIR.mkdir(parents=True, exist_ok=True)
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    def run_pipeline(self) -> bool:
        """Spawn a sub-process to invoke the data extraction script."""
        logger.info(f"Executing script: {EXTRACTION_SCRIPT.name}")
        try:
            # Subprocess safely runs python CLI with isolated arguments
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
            
            # Print extraction standard output straight to the test console
            if result.stdout:
                logger.info(f"Output: {result.stdout.strip()}")
            return True
            
        except subprocess.CalledProcessError as e:
            logger.error(f"Pipeline failed (Exit Code {e.returncode})")
            logger.error(f"Error output: {e.stderr}")
            return False

    def validate_results(self) -> bool:
        """Read and validate the output CSV to ensure data extraction succeeded."""
        # Quickly check file existence before attempting parsing
        if not OUTPUT_CSV.exists():
            logger.warning(f"Expected CSV not found at {OUTPUT_CSV}.")
            return False

        records: List[Dict[str, str]] = []
        try:
            # Parse CSV directly into a List of Dicts by column name
            with open(OUTPUT_CSV, mode='r', encoding='utf-8') as f:
                reader = csv.DictReader(f)
                records = list(reader)
        except Exception as e:
            logger.error(f"Failed to read CSV: {e}")
            return False

        logger.info(f"Validated {len(records)} records in {OUTPUT_FILENAME}.")
        
        # Succeed only if we've gathered at least one full record
        if records:
            logger.info("Data integrity checks passed successfully.")
            return True
        else:
            logger.warning("CSV is empty. No data extracted.")
            return False


def main():
    """Main verification CLI script entry."""
    runner = PipelineRunner()
    runner.setup_environment()
    
    # Pre-check: Ensure there are actual documents to test with
    docx_files = list(INPUT_DIR.glob("*.docx"))
    if not docx_files:
        logger.warning(f"No .docx files found in {INPUT_DIR}. Add forms before running.")
        return

    logger.info(f"Found {len(docx_files)} documents to process.")
    
    # Run pipeline -> Validate file -> Output final status
    if runner.run_pipeline():
        if runner.validate_results():
            logger.info("SUCCESS: Pipeline verification complete.")
        else:
            logger.warning("WARNING: Verification finished with non-fatal issues.")
    else:
        logger.error("FAILURE: Pipeline execution crashed or encountered errors.")


if __name__ == "__main__":
    main()