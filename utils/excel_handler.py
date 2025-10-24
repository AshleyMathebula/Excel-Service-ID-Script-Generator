"""
utils/excel_handler.py

Responsible for:
- Reading Excel workbook sheets
- Searching for Service_IDs
- Extracting number/code columns
- Cleaning numbers
- Formatting action lines

Key features:
- Uses pandas + openpyxl for reliable Excel parsing
- Handles 'service_id' lookup in a case-insensitive way
- Cleans and formats codes for downstream automation
"""

from typing import List, Iterable
import pandas as pd
from utils.logger import setup_logger  # Custom logger utility

# Initialize a logger specific to this module for tracking activity and errors
logger = setup_logger("excel_handler")


class ExcelHandler:
    """Encapsulate Excel reading and code processing logic."""

    def __init__(self, excel_path: str):
        """
        Constructor — initializes an ExcelHandler object.

        Args:
            excel_path (str): Path to the Excel file.

        Responsibilities:
        - Attempt to load the Excel file into memory.
        - Cache sheet names for quick access.
        - Log success or failure.

        Raises:
            FileNotFoundError: If the file cannot be opened or read.
        """
        self.path = excel_path  # Store the path for reference later
        try:
            # Use pandas.ExcelFile for efficient lazy sheet reading (does not load all sheets into memory)
            self._excel = pd.ExcelFile(self.path, engine="openpyxl")
            logger.info("Loaded Excel file: %s", self.path)
        except Exception as e:
            # Log and raise a descriptive error if the file cannot be opened
            logger.exception("Failed to open Excel file: %s", e)
            raise FileNotFoundError(f"Could not open Excel file: {e}")

    def sheet_names(self) -> List[str]:
        """
        Return a list of available sheet names in the workbook.

        Returns:
            List[str]: List of sheet names.
        """
        # ExcelFile already stores all sheet names after being opened
        return self._excel.sheet_names

    def find_service_codes(self, sheet_name: str, service_id: str) -> List[str]:
        """
        Search a specific sheet for rows where 'service_id' matches Service_<id> or <id>.

        Steps:
        1. Parse the sheet into a DataFrame.
        2. Normalize column names to lowercase.
        3. Locate the service_id column.
        4. Filter rows matching the given service.
        5. Extract 'sub-identifier' codes from matching rows.
        6. Return deduplicated list of codes.

        Args:
            sheet_name (str): Target sheet to search.
            service_id (str): The ID to match (e.g., '1056' or 'Service_1056').

        Returns:
            List[str]: Unique codes linked to the service.
        """
        try:
            # Parse the target sheet into a pandas DataFrame
            df = self._excel.parse(sheet_name)
        except Exception as e:
            # If the sheet cannot be read, log a warning and return an empty list
            logger.warning("Unable to read sheet %s: %s", sheet_name, e)
            return []

        # Normalize all column names (convert to lowercase and strip extra whitespace)
        df.columns = [str(c).lower().strip() for c in df.columns]

        # Try to identify which column holds the service IDs (e.g., "service_id")
        svc_col = next((c for c in df.columns if "service" in c and "id" in c), None)
        if not svc_col:
            # If no valid service ID column exists, log a warning and skip this sheet
            logger.warning("Sheet '%s' missing 'service_id' column.", sheet_name)
            return []

        # Ensure that the provided service_id has the proper prefix "service_"
        normalized_service = str(service_id).strip()
        if not normalized_service.lower().startswith("service_"):
            normalized_service = f"service_{normalized_service}"

        # Build a boolean mask to select all rows matching the target service_id
        mask = df[svc_col].astype(str).str.strip().str.lower() == normalized_service.lower()
        filtered = df.loc[mask]  # Subset DataFrame containing only matching rows

        # If no rows match, return an empty list
        if filtered.empty:
            return []

        # Try to identify which column holds the sub-identifiers / codes
        sub_col = next((c for c in df.columns if "sub" in c and "identifier" in c), None)
        if not sub_col:
            # If missing, log a warning and return nothing
            logger.warning("No 'sub-identifier' column found in sheet '%s'.", sheet_name)
            return []

        # Extract all codes from the 'sub-identifier' column
        # Filter out NaN or empty values
        codes = [str(v).strip() for v in filtered[sub_col] if pd.notna(v) and str(v).strip()]

        # Remove duplicates while preserving order using dict.fromkeys()
        unique_codes = list(dict.fromkeys(codes))

        # Log how many unique codes were found for this service and sheet
        logger.info(
            "Found %d raw code(s) for %s in sheet=%s",
            len(unique_codes),
            normalized_service,
            sheet_name,
        )

        return unique_codes  # Return final list of unique codes

    @staticmethod
    def clean_codes(codes: Iterable[str]) -> List[str]:
        """
        Clean and normalize raw codes.

        Steps:
        - Remove whitespace
        - Strip '?' characters
        - Keep '*' characters
        - Remove spaces and hyphens
        - Exclude empty results

        Args:
            codes (Iterable[str]): Raw codes to clean.

        Returns:
            List[str]: Cleaned, normalized codes.
        """
        cleaned = []  # Store cleaned results
        for raw in codes:
            s = str(raw).strip()  # Convert to string and remove leading/trailing spaces
            if not s:
                continue  # Skip empty entries

            # Remove unwanted characters but keep asterisks intact
            s = s.replace("?", "").replace(" ", "").replace("-", "")
            if s:
                cleaned.append(s)
        return cleaned

    @staticmethod
    def format_action_lines(codes: Iterable[str], username: str) -> List[str]:
        """
        Convert cleaned codes into formatted action lines.

        Example output:
            { "?.?.27840001402" }  : Actions SET_DEST_LA("cellfsc"),SET_ESME_GROUP(SAG_GROUP_1, A_ADDR)

        Args:
            codes (Iterable[str]): List of cleaned codes.
            username (str): Username to embed in the action line.

        Returns:
            List[str]: Formatted strings ready for output or script generation.
        """
        lines: List[str] = []  # Container for formatted action lines

        # Sanitize username (remove any quotes to avoid syntax errors in the output)
        username = username.replace('"', '').replace("'", "")

        # Build each formatted action line
        for code in codes:
            # Example pattern: { "?.?.<CODE>" }  : Actions ...
            line = f'{{ "?.?.{code}" }}  : Actions SET_DEST_LA("{username}"),SET_ESME_GROUP(SAG_GROUP_1, A_ADDR)'
            lines.append(line)

        return lines  # Return list of ready-to-write formatted lines
