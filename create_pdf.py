import os
import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from PyPDF2 import PdfReader, PdfWriter
import io
import re

def create_row_pdf(row_data, output_path):
    """
    Creates a PDF document from row data with formatted key-value pairs.
    
    Processes survey/form data by creating a PDF with each field as a labeled
    paragraph. URLs in the data are automatically converted to clickable blue links.
    Text content is wrapped to fit within page margins.
    
    Args:
        row_data (dict): Dictionary containing field names as keys and responses as values
        output_path (str): File path where the PDF should be saved
    
    Returns:
        None: Creates PDF file at specified output_path
    
    Note:
        - Filters out 'Matched Files' field from output
        - URLs are detected using regex pattern and made clickable
        - Long text content is automatically wrapped
    """
    doc = SimpleDocTemplate(output_path, pagesize=letter)
    story = []
    styles = getSampleStyleSheet()
    
    # Process data in chunks to avoid layout errors
    filtered_data = {k: v for k, v in row_data.items() if k != 'Matched Files'}
    
    for key, value in filtered_data.items():
        # Create key paragraph
        key_para = Paragraph(f"<b>{str(key)}:</b>", styles['Normal'])
        story.append(key_para)
        
        # Create value paragraph with wrapping and clickable URLs/files
        value_str = str(value) if pd.notna(value) else ''
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
    Merges a main PDF with multiple attachment PDFs into a single output file.
    
    Takes a primary PDF document and appends additional PDF files to create
    a consolidated document. Only processes files that exist and have .pdf extension.
    
    Args:
        main_pdf (str): Path to the primary PDF file to be merged first
        attachment_paths (list): List of file paths to PDF attachments to append
        output_path (str): File path where the merged PDF should be saved
    
    Returns:
        None: Creates merged PDF file at specified output_path
    
    Note:
        - Validates file existence and PDF extension before processing
        - Maintains page order: main PDF first, then attachments in list order
        - Silently skips invalid or missing attachment files
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

# Main execution
df = pd.read_excel('./outbox/updated_exported_data.xlsx', header=1)
os.makedirs('outbox/pdfs', exist_ok=True)

for idx, row in df.iterrows():
    # Create main PDF for row
    temp_pdf = f'outbox/temp_row_{idx}.pdf'
    create_row_pdf(row.to_dict(), temp_pdf)
    
    # Get matched files
    matched_files = str(row.get('Matched Files', '')).split('| ') if row.get('Matched Files') else []
    matched_files = [f.strip() for f in matched_files if f.strip()]
    
    # Get person's name from row 2, column R for filename
    person_name = str(row.iloc[17] if len(row) > 17 else f'row_{idx+1}').strip()
    # Sanitize filename by removing invalid characters
    safe_name = ''.join(c for c in person_name if c.isalnum() or c in (' ', '-', '_')).strip()
    final_pdf = f'outbox/pdfs/{safe_name}.pdf'
    
    if matched_files:
        merge_pdfs(temp_pdf, matched_files, final_pdf)
    else:
        os.rename(temp_pdf, final_pdf)
    
    # Clean up temp file
    if os.path.exists(temp_pdf):
        os.remove(temp_pdf)

print("✅ PDFs created in outbox/pdfs/")