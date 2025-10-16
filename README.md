## 🧮 Excel Service ID Script Generator

Interactive CLI tool to extract Service_IDs from Excel workbooks, preview record counts, and generate automated action scripts.

## 📚 Table of Contents

    Overview

    Features

    Installation

    Usage

    Example CLI Session

    Output Example

    Project Structure

    Algorithm & Data Structures

    Logging

    License

## 🧠 Overview

The Excel Service ID Script Generator is a Python CLI tool designed to automate the process of extracting service-related data from Excel workbooks and generating formatted action scripts.

It allows users to:

- Read Excel workbooks containing telecom or service configuration data.

- Select specific sheets or process all sheets at once.

- Input one or multiple Service_IDs and their associated usernames.

- Automatically extract and clean related codes or numbers.

- Preview record counts before generating scripts.

- Generate ready-to-use action script files in a structured output directory.

- Keep full activity logs for audit and traceability.

This tool is especially useful for telecom engineers, system administrators, and integration developers who handle batch configuration or routing scripts based on Service_IDs.

## ✨ Features

✅ Interactive CLI with dynamic sheet selection (all or specific sheets).
✅ Pre-generation summary of how many records exist per Service_ID in each sheet.
✅ Case-insensitive search for Service_<ID> patterns.
✅ Automatic code cleaning: removes ?, trims whitespace, and normalizes hyphens.
✅ Generates formatted action lines ready for system import.
✅ Organized output directory: output/.
✅ Centralized logging of all operations in logs/activity.log.
✅ Error handling for missing files, invalid sheet names, and malformed input.

## ⚙️ Installation

Clone the repository:

git clone https://github.com/your-username/excel-service-id-generator.git
cd excel-service-id-generator


Create a virtual environment (recommended):

python -m venv venv
source venv/bin/activate   # Linux/macOS
venv\Scripts\activate      # Windows


Install dependencies:

pip install -r requirements.txt

## 🧩 Dependencies

- pandas — Excel processing

- openpyxl — Excel engine

- pathlib — file path handling

- logging — centralized log management

## 🚀 Usage

Run the interactive CLI:

python main.py

Workflow

Enter the path to your Excel file (default: data/MO_Connection Database.xlsx)

Select sheets by number (e.g., 1,3) or type all

Enter one or more Service_IDs (comma-separated)

Provide a username (destination) for each Service_ID

The program counts matching records in each sheet

A summary table is displayed before generation

Confirm whether to proceed (y/n)

Scripts are generated in the output/ directory

💻 Example CLI Session
================================================================================
Excel Service ID Script Generator
================================================================================

Enter Excel file path (press Enter to use data/MO_Connection Database.xlsx):
Available sheets:
1. Billing_Data_Sep
2. Billing_Data_Oct
3. Archived_Logs

Select sheets by number (e.g., 1,3) or 'all': 1,2

Selected sheets: Billing_Data_Sep, Billing_Data_Oct
Enter one or more Service_IDs (comma-separated): 1056,2041
Enter username (destination) for Service_ID 1056: cellfsc
Enter username (destination) for Service_ID 2041: cellnew

## 📊 Scanning sheets for matching records...

  • Sheet: Billing_Data_Sep      | Service_1056 | User: cellfsc         | Codes found: 32
  • Sheet: Billing_Data_Sep      | Service_2041 | User: cellnew         | Codes found: 18
  • Sheet: Billing_Data_Oct      | Service_1056 | User: cellfsc         | Codes found: 40
  • Sheet: Billing_Data_Oct      | Service_2041 | User: cellnew         | Codes found: 25

================================================================================
Summary of results:
================================================================================
Sheet: Billing_Data_Sep          | Service_1056 | User: cellfsc         | Total: 32
Sheet: Billing_Data_Sep          | Service_2041 | User: cellnew         | Total: 18
Sheet: Billing_Data_Oct          | Service_1056 | User: cellfsc         | Total: 40
Sheet: Billing_Data_Oct          | Service_2041 | User: cellnew         | Total: 25
================================================================================
Proceed to generate scripts for these records? (y/n): y

✅ Generated 32 lines for Billing_Data_Sep → Billing_Data_Sep_1056_script.txt
✅ Generated 40 lines for Billing_Data_Oct → Billing_Data_Oct_1056_script.txt
✅ Generated 18 lines for Billing_Data_Sep → Billing_Data_Sep_2041_script.txt
✅ Generated 25 lines for Billing_Data_Oct → Billing_Data_Oct_2041_script.txt

🎉 Done! Processed 115 total records across selected sheets.

## 🧾 Output Example

Example output filename:

output/Billing_Data_Sep_1056_script.txt


Example action line:

{ "?.?.27840001402" }  : Actions SET_DEST_LA("cellfsc"),SET_ESME_GROUP(SAG_GROUP_1, A_ADDR)

## 🗂️ Project Structure
excel-service-id-generator/
│
├─ data/                       # Example Excel files
├─ logs/                       # Activity logs (activity.log)
├─ output/                     # Generated scripts
│
├─ utils/
│  ├─ excel_handler.py         # Excel reading, Service_ID search, code cleaning
│  ├─ file_writer.py           # File writing helper
│  └─ logger.py                # Centralized logging setup
│
├─ main.py                     # Interactive CLI entry point (Enhanced)
├─ requirements.txt            # Python dependencies
└─ README.md                   # Project documentation

## 🧩 Algorithm & Data Structures
Algorithm

Implements an ETL pipeline (Extract → Transform → Load):

Extract Excel rows by matching Service_ID.

Transform codes (clean, deduplicate, format).

Load results into structured .txt output files.

Employs Map–Filter–Reduce style processing for clarity and efficiency.

## Data Structures Used
Type	Purpose
list	Store filtered rows, cleaned codes, action lines
dict	Map Service_ID → username or service data
set	Remove duplicates
tuple	Temporary structured storage
str	Store identifiers and formatted output
pandas.DataFrame	Represent Excel sheets
pathlib.Path	Safe filesystem operations

Time Complexity: O(n) per sheet
Space Complexity: O(n) per sheet

## 🧾 Logging

All operations are logged both to the console and to a rotating log file:

logs/activity.log


Log entries include:

File loading and validation

Sheet selection and Service_ID parsing

Record counts and summary results

Output file creation

Errors or exceptions

Example:

[INFO] 2025-10-16 15:21:14 - Selected sheets: ['Billing_Data_Sep', 'Billing_Data_Oct']
[INFO] 2025-10-16 15:21:22 - Wrote 40 lines to output/Billing_Data_Oct_1056_script.txt

## 📜 License

MIT License © 2025
Developed by Ashley Mathebula (@Nika)
Feel free to use, modify, and distribute with attribution.