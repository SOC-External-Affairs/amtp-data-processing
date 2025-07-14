import os
import pandas as pd
from settings import REMATCH_INPUT_FILE, REMATCH_OUTPUT_FILE, REMATCH_SEARCH_ROOT, REMATCH_OUTPUT_DIR, REMATCH_EXCEL_ENGINE, REMATCH_MATCHED_FILES_COLUMN, REMATCH_FILE_SEPARATOR, REMATCH_RESPONSE_ID_KEYWORDS

# === CONFIGURATION ===
input_file = REMATCH_INPUT_FILE                      # Path to input Excel file containing response IDs
output_file = REMATCH_OUTPUT_FILE                    # Path to output Excel file that will contain matches
search_root = REMATCH_SEARCH_ROOT                    # Root directory to recursively search for matching files

# === LOAD EXCEL FILE ===
try:
    # Attempt to load Excel file using configured engine
    df = pd.read_excel(input_file, engine=REMATCH_EXCEL_ENGINE)
except FileNotFoundError:
    raise FileNotFoundError(f"Could not find the input file at path: {input_file}")

# === CHECK FOR RESPONSE ID COLUMN === 
# Searches for column containing configured keywords in name (case-insensitive)
response_id_col = None
for col in df.columns:
    if any(keyword in col.lower() for keyword in REMATCH_RESPONSE_ID_KEYWORDS):
        response_id_col = col
        break

if not response_id_col:
    print("Available columns:", list(df.columns))
    raise ValueError(f"No column containing {REMATCH_RESPONSE_ID_KEYWORDS} was found in the Excel file")

# === FUNCTION TO FIND MATCHING FILES ===
def find_matching_files(response_id, root_dir):
    """
    Recursively searches for files containing the response_id in their filename
    
    Args:
        response_id: ID to search for in filenames
        root_dir: Directory to start recursive search from
        
    Returns:
        list: Full paths of all matching files found
    """
    matches = []
    for dirpath, _, filenames in os.walk(root_dir):
        for filename in filenames:
            if str(response_id) in filename:
                full_path = os.path.join(dirpath, filename)
                matches.append(full_path)
    return matches

# === APPLY TO EACH ROW ===
# Create new column with separator-delimited list of matching file paths
df[REMATCH_MATCHED_FILES_COLUMN] = df[response_id_col].apply(
    lambda rid: REMATCH_FILE_SEPARATOR.join(find_matching_files(rid, search_root))
)

# === SAVE UPDATED FILE ===
# Create output directory if it doesn't exist and save Excel file
os.makedirs(REMATCH_OUTPUT_DIR, exist_ok=True)
df.to_excel(f'{REMATCH_OUTPUT_DIR}/{output_file}', index=False, engine=REMATCH_EXCEL_ENGINE)
print(f"✅ Updated file saved as: {output_file}")
