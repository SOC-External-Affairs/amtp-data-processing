# PDF Generation Settings

# Keywords to exclude from PDF output (using substring matching)
EXCLUDED_KEYWORDS = [
    'StartDate', 'EndDate', 'Status', 'IPAddress', 'Duration (in seconds)',
    'Finished', 'RecordedDate', 'RecipientLastName', 'RecipientFirstName',
    'RecipientEmail', 'ExternalReference', 'LocationLatitude', 'LocationLongitude',
    'DistributionChannel', 'Matched Files'
]

# File Processing Settings
INPUT_FILE = './inbox/data.xlsx'                     # Path to input Excel file containing response IDs
OUTPUT_FILE = 'updated_exported_data.xlsx'           # Path to output Excel file that will contain matches
SEARCH_ROOT = './inbox/'                             # Root directory to recursively search for matching files