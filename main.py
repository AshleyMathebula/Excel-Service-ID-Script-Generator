"""
main.py

Enhanced Version:
-----------------
Now includes a pre-generation summary step:
- After entering Service_IDs and usernames, the program counts matching codes
  in each selected sheet.
- Displays a table showing how many numbers were found per Service_ID per sheet.
- Asks the user for confirmation before generating output scripts.
"""

from pathlib import Path
from typing import List

from utils.excel_handler import ExcelHandler
from utils.file_writer import FileWriter
from utils.logger import setup_logger

# Configure centralized logger for this script
logger = setup_logger("main")


def display_menu(options: List[str]) -> None:
    """Pretty-print available sheet names as a numbered menu."""
    for i, opt in enumerate(options, start=1):
        print(f"{i}. {opt}")
    print()


def parse_indices(raw: str, max_index: int) -> List[int]:
    """Parse the user's sheet selection input ('all' or '1,3')."""
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
    """Prompt user to enter one or more Service_IDs (comma-separated)."""
    raw = input("Enter one or more Service_IDs (comma-separated): ").strip()
    return [s.strip() for s in raw.split(",") if s.strip()]


def main() -> None:
    """Main entry point for the CLI application."""
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
        # Step 2: Initialize helpers
        handler = ExcelHandler(str(excel_file))
        writer = FileWriter(output_dir="output")

        # Step 3: List available Excel sheets
        sheets = handler.sheet_names()
        if not sheets:
            logger.error("No sheets found in %s", excel_file)
            print("❌ No sheets found in the workbook.")
            return

        print("\nAvailable sheets:")
        display_menu(sheets)

        # Step 4: Select one or more sheets
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

        # Step 5: Ask for Service_IDs
        service_ids = input_service_ids()
        if not service_ids:
            print("❌ No Service_IDs provided. Exiting.")
            return

        # Step 6: For each Service_ID, ask for username
        service_usernames = {}
        for sid in service_ids:
            username = input(f"Enter username (destination) for Service_ID {sid}: ").strip()
            while not username:
                username = input("Username cannot be empty. Enter username: ").strip()
            service_usernames[sid] = username

        # Step 7: Preview counts before generation
        print("\n📊 Scanning sheets for matching records...\n")
        summary_table = []  # store (sheet, service_id, username, count)
        for sheet in selected_sheets:
            for sid, username in service_usernames.items():
                codes = handler.find_service_codes(sheet, sid)
                count = len(codes)
                summary_table.append((sheet, sid, username, count))
                print(f"  • Sheet: {sheet:<25} | Service_{sid:<6} | User: {username:<15} | Codes found: {count}")

        print("\n" + "=" * 80)
        print("Summary of results:")
        print("=" * 80)
        for sheet, sid, username, count in summary_table:
            print(f"Sheet: {sheet:<25} | Service_{sid:<6} | User: {username:<15} | Total: {count}")
        print("=" * 80)

        # Step 8: Ask for confirmation
        confirm = input("\nProceed to generate scripts for these records? (y/n): ").strip().lower()
        if confirm != "y":
            print("❌ Operation cancelled by user. No scripts generated.")
            logger.info("Operation cancelled after preview.")
            return

        # Step 9: Generate scripts
        total_processed = 0
        for sheet, sid, username, count in summary_table:
            if count == 0:
                continue
            codes = handler.find_service_codes(sheet, sid)
            cleaned = handler.clean_codes(codes)
            action_lines = handler.format_action_lines(cleaned, username)
            safe_sheet = sheet.replace(" ", "_")
            out_name = f"{safe_sheet}_{sid}_script.txt"
            out_path = Path("output") / out_name
            writer.write_text_file(out_path, action_lines)
            logger.info("Wrote %d lines to %s", len(action_lines), out_path)
            print(f"✅ Generated {len(action_lines)} lines for {sheet} → {out_name}")
            total_processed += len(action_lines)

        if total_processed == 0:
            print("\n⚠️ No matching records were processed.")
        else:
            print(f"\n🎉 Done! Processed {total_processed} total records across selected sheets.")
            logger.info("Completed. Total processed: %d", total_processed)

    except Exception as exc:
        logger.exception("Unhandled exception in main: %s", exc)
        print("❌ An unexpected error occurred. See logs/activity.log for details.")


if __name__ == "__main__":
    main()
