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
# SHARED SETTINGS
# ===================================================================

# Common Directory Paths
INBOX_PATH = './inbox'                              # Inbox directory path
INBOX_DATA_PATH = './inbox/data'                    # Inbox data directory path
OUTBOX_PATH = 'outbox'                              # Output directory path

# Common File Settings
EXCEL_ENGINE = 'openpyxl'                           # Excel engine to use for reading/writing Excel files
DATA_XLSX_FILENAME = 'data.xlsx'                    # Main data XLSX filename

# ===================================================================
# FILE PROCESSING SETTINGS (rematch.py)
# ===================================================================

# Input/Output File Paths
REMATCH_INPUT_FILE = f'{INBOX_PATH}/{DATA_XLSX_FILENAME}'  # Path to input Excel file containing response IDs
REMATCH_OUTPUT_FILE = 'updated_exported_data.xlsx'   # Path to output Excel file that will contain matches
REMATCH_SEARCH_ROOT = f'{INBOX_PATH}/'               # Root directory to recursively search for matching files
REMATCH_OUTPUT_DIR = OUTBOX_PATH                     # Directory name for output files

# Excel Processing Configuration
REMATCH_EXCEL_ENGINE = EXCEL_ENGINE                 # Excel engine to use for reading/writing Excel files
REMATCH_MATCHED_FILES_COLUMN = 'Matched Files'       # Column name for storing matched file paths
REMATCH_FILE_SEPARATOR = '| '                        # Separator for multiple file paths in matched files column

# Response ID Column Detection
REMATCH_RESPONSE_ID_KEYWORDS = ['response', 'id']    # Keywords to search for in column names to find Response ID column

# ===================================================================
# DOWNLOAD PROCESSING SETTINGS (process_downloads.py)
# ===================================================================

# Directory Paths
DOWNLOADS_INBOX_PATH = INBOX_PATH                    # Inbox directory path
DOWNLOADS_INBOX_DATA_PATH = INBOX_DATA_PATH          # Inbox data directory path

# File Processing Configuration
DOWNLOADS_TIME_WINDOW_HOURS = 2                     # Hours to look back for recent files
DOWNLOADS_FILE_PATTERN = '*AMTP*.zip'               # File pattern to match
DOWNLOADS_TARGET_XLSX = DATA_XLSX_FILENAME           # Target XLSX filename

