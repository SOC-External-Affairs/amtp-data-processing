import os
import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from PyPDF2 import PdfReader, PdfWriter
import io
import re
from settings import EXCLUDED_KEYWORDS

def create_row_pdf(row_data, output_path):
    """
    Creates a PDF document from survey row data with formatted key-value pairs.
    
    Processes Qualtrics survey data by creating a PDF with each field as a labeled
    paragraph. URLs and local file paths are automatically converted to clickable
    blue links. Text content wraps within page margins and preserves line breaks.
    
    Args:
        row_data (dict): Dictionary containing survey field names as keys and responses as values
        output_path (str): Absolute file path where the PDF should be saved
    
    Returns:
        None: Creates PDF file at specified output_path
    
    Features:
        - Filters out administrative fields using EXCLUDED_KEYWORDS from settings
        - Skips fields with empty or whitespace-only values
        - Converts URLs to clickable blue links (https://...)
        - Converts local file paths to clickable blue links (./...)
        - Preserves line breaks by converting \n to HTML <br/> tags
        - Uses ReportLab's Paragraph elements for automatic text wrapping
    """
    doc = SimpleDocTemplate(output_path, pagesize=letter)
    story = []
    styles = getSampleStyleSheet()
    
    # Load excluded keywords from settings
    excluded_keywords = EXCLUDED_KEYWORDS
    
    # Process data in chunks to avoid layout errors
    filtered_data = {}
    for k, v in row_data.items():
        # Check if any excluded keyword is in the header
        if not any(keyword in k for keyword in excluded_keywords):
            filtered_data[k] = v
    
    for key, value in filtered_data.items():
        # Skip if value is empty or only whitespace
        value_str = str(value) if pd.notna(value) else ''
        if not value_str.strip():
            continue
            
        # Create key paragraph
        key_para = Paragraph(f"<b>{str(key)}:</b>", styles['Normal'])
        story.append(key_para)
        
        # Create value paragraph with wrapping and clickable URLs/files
        # Preserve line breaks by converting to HTML br tags
        value_str = value_str.replace('\n', '<br/>')
       
        # Make URLs clickable and blue
        url_pattern = r'(https?://[^\s]+)'
        value_str = re.sub(url_pattern, r'<link href="\1" color="blue">\1</link>', value_str)
       
        # Make local file paths clickable and blue
        file_pattern = r'(\./[^\n\r]+)'
        value_str = re.sub(file_pattern, r'<link href="file://\1" color="blue">\1</link>', value_str)
        value_para = Paragraph(value_str, styles['Normal'])
        story.append(value_para)
        story.append(Spacer(1, 12))
    
    doc.build(story)

def merge_pdfs(main_pdf, attachment_paths, output_path):
    """
    Merges a survey response PDF with associated attachment PDFs into a single document.
    
    Takes a primary survey response PDF and appends matched attachment files to create
    a consolidated document. Used to combine survey responses with uploaded files
    that were matched by Response ID.
    
    Args:
        main_pdf (str): Absolute path to the primary survey response PDF
        attachment_paths (list): List of absolute file paths to PDF attachments to append
        output_path (str): Absolute file path where the merged PDF should be saved
    
    Returns:
        None: Creates merged PDF file at specified output_path
    
    Behavior:
        - Validates file existence and .pdf extension before processing
        - Maintains page order: main survey PDF first, then attachments in list order
        - Silently skips invalid, missing, or non-PDF attachment files
        - Uses PyPDF2 for reliable PDF merging operations
        - Overwrites existing files at output_path
    """
    writer = PdfWriter()
    
    # Add main PDF
    with open(main_pdf, 'rb') as f:
        reader = PdfReader(f)
        for page in reader.pages:
            writer.add_page(page)
    
    # Add attachments
    for path in attachment_paths:
        if os.path.exists(path) and path.lower().endswith('.pdf'):
            with open(path, 'rb') as f:
                reader = PdfReader(f)
                for page in reader.pages:
                    writer.add_page(page)
    
    with open(output_path, 'wb') as f:
        writer.write(f)

# === MAIN EXECUTION ===
# Process Qualtrics export with merged headers from rows 0 and 1
# Read Excel without header to access both header rows
df_raw = pd.read_excel('./outbox/updated_exported_data.xlsx', header=None)

# Combine row 0 and row 1 to create merged headers for Qualtrics format
# Row 0: Question labels, Row 1: Question text (trimmed to first 3 words)
headers = []
for col in range(len(df_raw.columns)):
    row0_val = str(df_raw.iloc[0, col]).strip() if pd.notna(df_raw.iloc[0, col]) else ''
    row1_val = str(df_raw.iloc[1, col]).strip() if pd.notna(df_raw.iloc[1, col]) else ''
    combined = f"{row0_val} {row1_val}".strip()
    headers.append(combined)

# Create dataframe with merged headers and actual survey response data from row 2 onwards
df = pd.DataFrame(df_raw.iloc[2:].values, columns=headers)
df.reset_index(drop=True, inplace=True)

os.makedirs('outbox/pdfs', exist_ok=True)

# Process each survey response row to create individual PDFs
for idx, row in df.iterrows():
    # Create main survey response PDF for current row
    temp_pdf = f'outbox/temp_row_{idx}.pdf'
    create_row_pdf(row.to_dict(), temp_pdf)
    
    # Get matched attachment files from 'Matched Files' column (populated by rematch.py)
    if 'Matched Files' in df.columns:
        matched_files_col_idx = df.columns.get_loc('Matched Files')
        matched_files_col = row.iloc[matched_files_col_idx] if len(row) > matched_files_col_idx else ''
        matched_files = str(matched_files_col).split('| ') if pd.notna(matched_files_col) and matched_files_col else []
    else:
        matched_files = []    
    matched_files = [f.strip() for f in matched_files if f.strip()]
    
    # Extract person's name (column R, index 17) and show name (column Z, index 25) for filename
    person_name = str(row.iloc[17] if len(row) > 17 else f'row_{idx+1}').strip()
    show_name = str(row.iloc[25] if len(row) > 25 else '').strip()  # Column Z is index 25
    
    # Create descriptive filename: "Show Name - Person Name.pdf" or "Person Name.pdf"
    if show_name:
        filename_parts = [show_name, person_name]
    else:
        filename_parts = [person_name]
    
    # Sanitize filename by removing filesystem-invalid characters
    safe_name = ' - '.join(filename_parts)
    safe_name = ''.join(c for c in safe_name if c.isalnum() or c in (' ', '-', '_')).strip()
    final_pdf = f'outbox/pdfs/{safe_name}.pdf'
    
    # Merge survey response with matched attachment files or save standalone PDF
    print(f"Processing {safe_name}: {len(matched_files)} matched files")
    if matched_files:
        print(f"  Matched files: {matched_files}")
        merge_pdfs(temp_pdf, matched_files, final_pdf)
        # Clean up temporary file after successful merge
        if os.path.exists(temp_pdf):
            os.remove(temp_pdf)
    else:
        # No attachments found, rename temp file to final name
        os.rename(temp_pdf, final_pdf)

print("✅ PDFs created in outbox/pdfs/")