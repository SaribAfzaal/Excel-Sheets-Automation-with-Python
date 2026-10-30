# Conversation Log

## User

At first, save this conversation in a separate file named claude.md

Build a Python script that automates common Excel tasks across multiple files.

Requirements:
- Accept multiple Excel/CSV files from a folder and merge them into a single clean Excel file, handling mismatched columns gracefully
- Remove duplicate rows and handle missing values (configurable: drop, fill with default, or flag)
- Apply basic formatting: bold headers, auto-fit column widths, alternating row colors
- Add conditional formatting for a specified column (e.g., highlight values above/below a threshold)
- Generate a summary sheet with totals/averages for numeric columns
- Use pandas for data handling and openpyxl for formatting
- Make it configurable via a simple config file or command-line arguments (input folder, output file, column to apply conditional formatting on)
- Add a requirements.txt file and README with usage examples

Also generate a small sample/dummy dataset (e.g., sample_data.xlsx or sample_input.pdf) so I can test the script immediately after building it.

## Assistant

Saved this conversation to claude.md, then building the project:
- `excel_automation.py` (main script)
- `config.ini` (sample config)
- `generate_sample_data.py` + `sample_input/` (dummy dataset)
- `requirements.txt`, `README.md`
