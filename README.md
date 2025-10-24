## 🧮 Excel Service ID Script Generator

A powerful interactive Python CLI tool that automates the extraction of Service_ID data from Excel workbooks and generates ready-to-use action scripts.

Ideal for telecom engineers, system administrators, and integration developers handling batch configuration or routing scripts.

## 📚 Table of Contents

    Overview
    
    Features
    
    Installation
    
    Dependencies
    
    Usage
    
    Example CLI Session
    
    Output Example
    
    Project Structure
    
    Algorithm & Data Structures
    
    Logging
    
    License

## 🧠 Overview

The Excel Service ID Script Generator simplifies the process of scanning Excel workbooks, finding specific Service_ID values, and generating formatted configuration scripts.

It allows users to:

Read Excel workbooks containing service or telecom configuration data.

Select one, multiple, or all sheets dynamically.

Enter one or more Service_IDs and associate them with usernames.

Clean and normalize service codes automatically.

Preview record counts per Service_ID before generating scripts.

Generate formatted script files for import or deployment.

Keep detailed activity logs for traceability and audits.

## ✨ Features

✅ Interactive CLI with sheet selection (all or specific by number).
✅ Pre-generation summary showing record counts per Service_ID.
✅ Case-insensitive matching for Service_<ID> patterns.
✅ Automatic data cleaning (removes ?, trims spaces, fixes hyphens).
✅ Organized output directory structure (output/).
✅ Centralized logging in logs/activity.log.
✅ Graceful error handling for missing files, invalid sheets, or malformed input.

## ⚙️ Installation
1. Clone the repository:
git clone https://github.com/your-username/excel-service-id-generator.git
cd excel-service-id-generator

2. Create and activate a virtual environment:
python -m venv venv
Linux/macOS
source venv/bin/activate
Windows
venv\Scripts\activate

3. Install dependencies:
pip install -r requirements.txt

## 🧩 Dependencies

pandas — Excel processing and data handling

openpyxl — Excel engine

pathlib — Path management

logging — Activity tracking and debugging

## 🚀 Usage

Run the tool interactively:

python main.py

Workflow:

Enter the path to your Excel file (default: data/MO_Connection Database.xlsx)

Select sheets by number (e.g., 1,3) or type all

Enter one or more Service_IDs (comma-separated)

Provide usernames for each Service_ID

Preview record counts for each sheet and ID

Confirm to generate scripts

Scripts are saved in the output/ directory

## 💻 Example CLI Session
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

📊 Scanning sheets for matching records...

• Sheet: Billing_Data_Sep | Service_1056 | User: cellfsc | Codes found: 32  
• Sheet: Billing_Data_Sep | Service_2041 | User: cellnew | Codes found: 18  
• Sheet: Billing_Data_Oct | Service_1056 | User: cellfsc | Codes found: 40  
• Sheet: Billing_Data_Oct | Service_2041 | User: cellnew | Codes found: 25  

================================================================================
Summary of results:
================================================================================
Sheet: Billing_Data_Sep | Service_1056 | User: cellfsc | Total: 32  
Sheet: Billing_Data_Sep | Service_2041 | User: cellnew | Total: 18  
Sheet: Billing_Data_Oct | Service_1056 | User: cellfsc | Total: 40  
Sheet: Billing_Data_Oct | Service_2041 | User: cellnew | Total: 25  
================================================================================
Proceed to generate scripts for these records? (y/n): y

✅ Generated 32 lines → Billing_Data_Sep_1056_script.txt  
✅ Generated 40 lines → Billing_Data_Oct_1056_script.txt  
✅ Generated 18 lines → Billing_Data_Sep_2041_script.txt  
✅ Generated 25 lines → Billing_Data_Oct_2041_script.txt  

🎉 Done! Processed 115 total records across selected sheets.

## 🧾 Output Example

Output Directory:

output/
  ├─ Billing_Data_Sep_1056_script.txt
  ├─ Billing_Data_Oct_1056_script.txt


Example Line:

{ "?.?.27840001402" } : Actions SET_DEST_LA("cellfsc"), SET_ESME_GROUP(SAG_GROUP_1, A_ADDR)

## 🗂️ Project Structure
excel-service-id-generator/
│
├─ data/                       # Example Excel files
├─ logs/                       # Activity logs (activity.log)
├─ output/                     # Generated scripts
│
├─ utils/
│  ├─ excel_handler.py         # Excel reading, ID search, and code cleaning
│  ├─ file_writer.py           # File writing helper
│  └─ logger.py                # Centralized logging setup
│
├─ main.py                     # Interactive CLI entry point
├─ requirements.txt            # Python dependencies
└─ README.md                   # Project documentation

## 🧩 Algorithm & Data Structures
Algorithm

Implements a simple ETL pipeline:
Extract → Transform → Load

Extract: Filter Excel rows by matching Service_ID.

Transform: Clean, normalize, and deduplicate codes.

Load: Generate formatted .txt output files.

Data Structures
Type	Purpose
list	Store filtered rows and formatted lines
dict	Map Service_ID → username or records
set	Remove duplicate codes
tuple	Temporary structured data
pandas.DataFrame	Represent Excel sheets
pathlib.Path	Safe and portable file operations

Time Complexity: O(n) per sheet
Space Complexity: O(n) per sheet

## 🧾 Logging

Logs are written both to the console and to:

logs/activity.log


Log Includes:

Excel file loading

Sheet and Service_ID selections

Record counts per sheet

Output file creation

Errors and exceptions

Example:

[INFO] 2025-10-16 15:21:14 - Selected sheets: ['Billing_Data_Sep', 'Billing_Data_Oct']
[INFO] 2025-10-16 15:21:22 - Wrote 40 lines to output/Billing_Data_Oct_1056_script.txt

## 📜 License

MIT License © 2025
Developed by Ashley Mathebula (@Nika)

Feel free to use, modify, and distribute with proper attribution.
