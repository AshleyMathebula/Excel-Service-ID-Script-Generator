"""
service_summary.py

This module handles generation of per-service summary text files that
aggregate script outputs from multiple Excel sheets. Each generated file
contains all related records for a given Service_ID and its assigned user.
"""

from pathlib import Path


def generate_service_files(summary_table, service_usernames, script_contents, output_dir="output"):
    """
    Generate per-service summary files combining script data across multiple sheets.

    This function consolidates generated script contents for each Service_ID into a
    single summary file. It ensures that each Service_ID’s final report contains
    the actual script lines from all associated sheets.

    Args:
        summary_table (list[tuple]):
            List of tuples in the format (sheet, service_id, username, count),
            summarizing data collected from Excel processing.
        service_usernames (dict):
            Mapping of {service_id: username}, associating each Service_ID
            with its destination username.
        script_contents (dict):
            Mapping of {(sheet, service_id): [lines]} where each value is the
            list of script lines previously generated per sheet.
        output_dir (str, optional):
            Path to the output directory where summary files will be saved.
            Defaults to "output".

    Returns:
        None. Writes one summary file per Service_ID to disk.

    Example:
        generate_service_files(summary_table, service_usernames, script_contents)
    """
    # -----------------------------------------------------------------------
    # Step 1: Prepare output directory
    # -----------------------------------------------------------------------
    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)

    # -----------------------------------------------------------------------
    # Step 2: Group sheet names by Service_ID
    # -----------------------------------------------------------------------
    # Build a dictionary like:
    # {
    #   "12345": ["Sheet1", "Sheet3"],
    #   "67890": ["Sheet2"]
    # }
    # Only include Service_IDs that have at least one record (count > 0)
    service_dict = {}
    for sheet, sid, username, count in summary_table:
        if count == 0:
            continue
        service_dict.setdefault(sid, []).append(sheet)

    # -----------------------------------------------------------------------
    # Step 3: Generate one summary file per Service_ID
    # -----------------------------------------------------------------------
    for sid, sheets in service_dict.items():
        username = service_usernames[sid]
        file_name = f"{username}_{sid}_summary.txt"
        file_path = output_path / file_name

        # Open output file for writing in UTF-8 encoding
        with open(file_path, "w", encoding="utf-8") as f:
            # Write descriptive header section
            f.write(f"{username.upper()}\n\n")
            f.write(f"{username:<12} service_{sid}\n\n")
            f.write("OA: \n\n")

            # Iterate through all sheets related to this Service_ID
            for i, sheet in enumerate(sheets):
                safe_sheet = sheet.replace(" ", "_")
                f.write(f"{safe_sheet}_script\n")

                # Retrieve previously generated script content
                content = script_contents.get((sheet, sid), [])
                if content:
                    f.write("\n".join(content) + "\n")

                # Add separator line between sheet blocks (except last one)
                if i < len(sheets) - 1:
                    f.write("\n*****************************************\n\n")

        # Log output for user visibility
        print(f"[INFO] Summary for Service_{sid} written to: {file_path}")
