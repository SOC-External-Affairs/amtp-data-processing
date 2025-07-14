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

# Run the scripts
python process_downloads.py
python rematch.py
python create_pdf.py