# Automated PDF Generation System for Qualtrics Survey Data

This system processes Qualtrics survey exports for AMTP submissions by automatically generating individual PDFs that include:

- Respondent information
- Survey responses 
- All uploaded attachments

The output is a single consolidated PDF per submission, named with the respondent's name and play title, containing the complete submission package.## Overview

This pipeline automatically processes AMTP survey data by:
1. **Extracting** zip files from Downloads folder
2. **Matching** survey responses with uploaded files
3. **Generating** individual PDFs for each respondent with their attachments

## Quick Start

```bash
./run.sh
```

That's it! The script handles everything automatically.

## What It Does

### 1. Download Processing (`process_downloads.py`)
- Scans `~/Downloads/` for recent AMTP zip files (last 2 hours)
- Extracts survey data (XLSX) to `./inbox/data.xlsx`
- Extracts all attachments to `./inbox/data/`

### 2. File Matching (`rematch.py`)
- Reads survey responses from `./inbox/data.xlsx`
- Matches uploaded files by Response ID
- Creates `./outbox/updated_exported_data.xlsx` with file paths

### 3. PDF Generation (`create_pdf.py`)
- Creates individual PDFs for each survey response
- Names files: `"Show Name - Person Name.pdf"`
- Includes clickable links for URLs and file paths
- Merges matched attachments into each PDF
- Outputs to `./outbox/pdfs/`

## File Structure

```
amtp-analysis/
├── run.sh                    # Main execution script
├── process_downloads.py      # Download processing
├── rematch.py               # File matching
├── create_pdf.py            # PDF generation
├── settings.py              # Configuration
├── inbox/
│   ├── data.xlsx           # Survey data
│   └── data/               # Uploaded files
└── outbox/
    ├── updated_exported_data.xlsx
    └── pdfs/               # Generated PDFs
```

## Configuration

Edit `settings.py` to customize:
- File paths and directories
- Time windows for file processing
- PDF exclusion keywords
- Excel processing options

## Requirements

- Python 3.7+
- macOS/Linux (for bash script)
- Dependencies installed automatically:
  - pandas
  - openpyxl
  - reportlab
  - PyPDF2

## Features

- **Automatic Setup**: Creates virtual environment and installs dependencies
- **Smart File Detection**: Finds recent AMTP files automatically
- **Response Matching**: Links survey responses to uploaded files
- **Rich PDFs**: Clickable URLs, preserved formatting, merged attachments
- **Error Handling**: Graceful handling of missing files and data
- **Configurable**: All settings centralized in `settings.py`

## Troubleshooting

**No files found?**
- Check Downloads folder for AMTP zip files
- Verify files are less than 2 hours old
- Check file naming contains "AMTP"

**PDFs missing attachments?**
- Ensure Response IDs match between survey and file names
- Check `./inbox/data/` for extracted files
- Review console output for matching details

**Permission errors?**
- Make run.sh executable: `chmod +x run.sh`
- Check file permissions in Downloads folder