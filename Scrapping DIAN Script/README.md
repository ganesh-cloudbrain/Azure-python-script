
# Automated Invoice Processing System

## Overview
This Python script automates the process of downloading, processing, and organizing invoices. It interacts with web pages, emails, and XML files to extract and process invoice data. Key functions include:
- Fetching URLs from emails.
- Downloading and unzipping invoice files.
- Parsing XML files to extract invoice data.
- Writing extracted data to Excel files.
- Organizing downloaded files and logs.

## Requirements
- Python 3.x
- `google-auth-oauthlib` (For Google OAuth)
- `google-auth` (For Google authentication)
- `googleapiclient` (For accessing Google APIs)
- `playwright` (For browser automation)
- `beautifulsoup4` (For HTML and XML parsing)
- `requests` (For making HTTP requests)
- `openpyxl` (For Excel file operations)
- `shutil` and `os` (For file and directory operations)
- `re` (For regular expressions)
- `datetime` (For date and time operations)
- `base64` (For encoding operations)
- `html` (For HTML entities and character manipulation)

Install the dependencies using pip:
```bash
pip install google-auth-oauthlib google-auth google-api-python-client playwright beautifulsoup4 requests openpyxl
```

## Usage Instructions
1. Clone or download this script to your local machine.
2. Ensure you have all the required dependencies installed.
3. Run the script using Python:
   ```bash
   python main.py
   ```

## Input
- User and Company codes in a text file.
- Date ranges for invoice processing.
- Email credentials and setup for receiving URLs.

## Output
- Downloaded and unzipped invoice files.
- Excel files with extracted invoice data.
- Organized files in respective directories.
- Log files with details of script execution.
