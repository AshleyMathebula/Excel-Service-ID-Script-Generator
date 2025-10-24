"""
main.py

Enhanced Version (with Service File Generation)
-----------------------------------------------
This script provides a CLI utility that processes an Excel workbook containing
Service_ID records, generates individual scripts per Service_ID, and then
creates summary files combining these results.

Key Features:
- Allows user to select specific or all sheets in an Excel workbook.
- Accepts one or more Service_IDs with corresponding usernames.
- Displays a summary table showing the count of found records before generation.
- Prompts for user confirmation before file creation.
- Generates both per-sheet scripts and consolidated service summary files.

Dependencies:
- utils.excel_handler.ExcelHandler: Handles Excel operations.
- utils.file_writer.FileWriter: Handles text file output.
- utils.logger.setup_logger: Provides consistent logging configuration.
- utils.service_summary.generate_service_files: Builds final per-service summaries.
"""

from pathlib import Path
from typing import List

from utils.excel_handler import ExcelHandler
from utils.file_writer import FileWriter
from utils.logger import setup_logger
from utils.service_summary import generate_service_files

# ---------------------------------------------------------------------------
# Global logger configuration
# ---------------------------------------------------------------------------
# Initializes a centralized logger instance for the script to record events
logger = setup_logger("main")


# ---------------------------------------------------------------------------
# Helper Functions
# ---------------------------------------------------------------------------

def display_menu(options: List[str]) -> None:
    """
    Displays a numbered list of options (e.g., Excel sheet names).

    Args:
        options (List[str]): List of available options to display.
    """
    for i, opt in enumerate(options, start=1):
        print(f"{i}. {opt}")
    print()


def parse_indices(raw: str, max_index: int) -> List[int]:
    """
    Parses a user's input string for sheet selection.

    Supports either:
      - 'all' keyword to select all sheets
      - Comma-separated indices like '1,3,4'

    Args:
        raw (str): Raw input from the user.
        max_index (int): Total number of available sheets.

    Returns:
        List[int]: Zero-based indices of selected sheets.

    Raises:
        ValueError: If the input contains invalid characters.
        IndexError: If an index is out of valid range.
    """
    raw = raw.strip().lower()
    if raw == "all":
        return list(range(max_index))

    parts = [p.strip() for p in raw.split(",") if p.strip()]
    indices: List[int] = []

    for p in parts:
        if not p.isdigit():
            raise ValueError("Sheet selections must be numbers or 'all'.")
        idx = int(p) - 1
        if idx < 0 or idx >= max_index:
            raise IndexError("Sheet index out of range.")
        indices.append(idx)
    return indices


def input_service_ids() -> List[str]:
    """
    Prompts the user to enter one or more Service_IDs.

    Returns:
        List[str]: List of trimmed Service_ID strings.
    """
    raw = input("Enter one or more Service_IDs (comma-separated): ").strip()
    return [s.strip() for s in raw.split(",") if s.strip()]


# ---------------------------------------------------------------------------
# Main Application Logic
# ---------------------------------------------------------------------------

def main() -> None:
    """
    Entry point for the Excel Service ID Script Generator.

    Handles user interaction, data extraction from Excel, record validation,
    and output file generation. Includes error handling, preview confirmation,
    and logging for audit purposes.
    """
    logger.info("Excel Service ID Script Generator started")

    print("=" * 80)
    print("Excel Service ID Script Generator (Enhanced Version)")
    print("=" * 80)

    # Step 1: Ask for Excel file path or use default
    excel_path = input(
        "Enter Excel file path (press Enter to use data/MO_Connection Database.xlsx): "
    ).strip()
    if not excel_path:
        excel_path = "data/MO_Connection Database.xlsx"

    excel_file = Path(excel_path)
    if not excel_file.exists():
        logger.error("Excel file not found: %s", excel_file)
        print(f"❌ Excel file not found: {excel_file}")
        return

    try:
        # Step 2: Initialize helper classes for Excel I/O and file writing
        handler = ExcelHandler(str(excel_file))
        writer = FileWriter(output_dir="output")

        # Step 3: List all available sheets in the workbook
        sheets = handler.sheet_names()
        if not sheets:
            logger.error("No sheets found in %s", excel_file)
            print("❌ No sheets found in the workbook.")
            return

        print("\nAvailable sheets:")
        display_menu(sheets)

        # Step 4: Allow user to select one or more sheets by index or 'all'
        while True:
            selection = input("Select sheets by number (e.g., 1,3) or 'all': ").strip()
            try:
                selected_indices = parse_indices(selection, len(sheets))
                selected_sheets = [sheets[i] for i in selected_indices]
                break
            except (ValueError, IndexError) as e:
                print(f"⚠️ {e} Please try again.")

        print(f"\nSelected sheets: {', '.join(selected_sheets)}")
        logger.info("Selected sheets: %s", selected_sheets)

        # Step 5: Request Service_IDs from user
        service_ids = input_service_ids()
        if not service_ids:
            print("❌ No Service_IDs provided. Exiting.")
            return

        # Step 6: Request username for each Service_ID
        service_usernames = {}
        for sid in service_ids:
            username = input(f"Enter username (destination) for Service_ID {sid}: ").strip()
            while not username:
                username = input("Username cannot be empty. Enter username: ").strip()
            service_usernames[sid] = username

        # Step 7: Display record preview before generation
        print("\n📊 Scanning sheets for matching records...\n")
        summary_table = []  # List of tuples (sheet, service_id, username, count)

        for sheet in selected_sheets:
            for sid, username in service_usernames.items():
                codes = handler.find_service_codes(sheet, sid)
                count = len(codes)
                summary_table.append((sheet, sid, username, count))
                print(f"  • Sheet: {sheet:<25} | Service_{sid:<6} | User: {username:<15} | Codes found: {count}")

        # Display summary results in a clean table
        print("\n" + "=" * 80)
        print("Summary of results:")
        print("=" * 80)
        for sheet, sid, username, count in summary_table:
            print(f"Sheet: {sheet:<25} | Service_{sid:<6} | User: {username:<15} | Total: {count}")
        print("=" * 80)

        # Step 8: Ask user to confirm before generating scripts
        confirm = input("\nProceed to generate scripts for these records? (y/n): ").strip().lower()
        if confirm != "y":
            print("❌ Operation cancelled by user. No scripts generated.")
            logger.info("Operation cancelled after preview.")
            return

        # Step 9: Generate scripts and keep track of processed lines
        script_contents = {}  # Dict mapping (sheet, sid) → List[str] of script lines
        total_processed = 0

        for sheet, sid, username, count in summary_table:
            if count == 0:
                continue  # Skip empty results

            # Retrieve, clean, and format service codes
            codes = handler.find_service_codes(sheet, sid)
            cleaned = handler.clean_codes(codes)
            action_lines = handler.format_action_lines(cleaned, username)

            # Construct safe output filename and write to disk
            safe_sheet = sheet.replace(" ", "_")
            out_name = f"{safe_sheet}_{sid}_script.txt"
            out_path = Path("output") / out_name
            writer.write_text_file(out_path, action_lines)

            # Track script content for later summary file generation
            script_contents[(sheet, sid)] = action_lines
            logger.info("Wrote %d lines to %s", len(action_lines), out_path)
            print(f"✅ Generated {len(action_lines)} lines for {sheet} → {out_name}")
            total_processed += len(action_lines)

        # Step 10: Generate per-service summary files combining all results
        print("\n📦 Generating consolidated service summary files...\n")
        generate_service_files(summary_table, service_usernames, script_contents, output_dir="output")
        logger.info("Service summary files generated successfully.")

        # Step 11: Final completion message
        if total_processed == 0:
            print("\n⚠️ No matching records were processed.")
        else:
            print(f"\n🎉 Done! Processed {total_processed} total records across selected sheets.")
            logger.info("Completed. Total processed: %d", total_processed)

    except Exception as exc:
        # Catch and log any unhandled runtime exception
        logger.exception("Unhandled exception in main: %s", exc)
        print("❌ An unexpected error occurred. See logs/activity.log for details.")


# ---------------------------------------------------------------------------
# Script Entry Point
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    main()
