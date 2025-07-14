#!/bin/bash

# Create venv if it doesn't exist
if [ ! -d "venv" ]; then
    echo "Creating virtual environment..."
    python3 -m venv venv
fi

# Activate venv
source venv/bin/activate

# Upgrade pip if needed
pip install --upgrade pip -q

# Create requirements.txt if it doesn't exist
if [ ! -f "requirements.txt" ]; then
    echo "pandas" > requirements.txt
    echo "openpyxl" >> requirements.txt
    echo "reportlab" >> requirements.txt
    echo "PyPDF2" >> requirements.txt
fi

# Install requirements if not already installed
pip install -q -r requirements.txt



# move and extract zip files from downloads folder
python process_downloads.py

# pair up the responses with uploaded PDF files that much
python rematch.py

# Generate final PDF report from the processed data
python create_pdf.py
