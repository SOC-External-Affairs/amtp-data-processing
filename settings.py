# ===================================================================
# PDF GENERATION SETTINGS
# ===================================================================

# Keywords to exclude from PDF output (using substring matching)
# These Qualtrics administrative fields will be filtered out of final PDFs
PDF_EXCLUDED_KEYWORDS = [
    'StartDate', 'EndDate', 'Status', 'IPAddress', 'Duration (in seconds)',
    'Finished', 'RecordedDate', 'RecipientLastName', 'RecipientFirstName',
    'RecipientEmail', 'ExternalReference', 'LocationLatitude', 'LocationLongitude',
    'DistributionChannel', 
]

# ===================================================================
# FILE PROCESSING SETTINGS (rematch.py)
# ===================================================================

# Input/Output File Paths
REMATCH_INPUT_FILE = './inbox/data.xlsx'             # Path to input Excel file containing response IDs
REMATCH_OUTPUT_FILE = 'updated_exported_data.xlsx'   # Path to output Excel file that will contain matches
REMATCH_SEARCH_ROOT = './inbox/'                     # Root directory to recursively search for matching files
REMATCH_OUTPUT_DIR = 'outbox'                        # Directory name for output files

# Excel Processing Configuration
REMATCH_EXCEL_ENGINE = 'openpyxl'                   # Excel engine to use for reading/writing Excel files
REMATCH_MATCHED_FILES_COLUMN = 'Matched Files'       # Column name for storing matched file paths
REMATCH_FILE_SEPARATOR = '| '                        # Separator for multiple file paths in matched files column

# Response ID Column Detection
REMATCH_RESPONSE_ID_KEYWORDS = ['response', 'id']    # Keywords to search for in column names to find Response ID column

