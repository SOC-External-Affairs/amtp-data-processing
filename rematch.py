import os
import pandas as pd
from settings import INPUT_FILE, OUTPUT_FILE, SEARCH_ROOT

# === CONFIGURATION ===
input_file = INPUT_FILE                              # Path to input Excel file containing response IDs
output_file = OUTPUT_FILE                            # Path to output Excel file that will contain matches
search_root = SEARCH_ROOT                            # Root directory to recursively search for matching files

# === LOAD EXCEL FILE ===
try:
    # Attempt to load Excel file using openpyxl engine
    df = pd.read_excel(input_file, engine='openpyxl')
except FileNotFoundError:
    raise FileNotFoundError(f"Could not find the input file at path: {input_file}")

# === CHECK FOR RESPONSE ID COLUMN === 
# Searches for column containing 'response' or 'id' in name (case-insensitive)
response_id_col = None
for col in df.columns:
    if 'response' in col.lower() or 'id' in col.lower():
        response_id_col = col
        break

if not response_id_col:
    print("Available columns:", list(df.columns))
    raise ValueError("No column containing 'response' or 'id' was found in the Excel file")

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
# Create new column with pipe-separated list of matching file paths
df['Matched Files'] = df[response_id_col].apply(
    lambda rid: '| '.join(find_matching_files(rid, search_root))
)

# === SAVE UPDATED FILE ===
# Create outbox directory if it doesn't exist and save Excel file
os.makedirs('outbox', exist_ok=True)
df.to_excel(f'outbox/{output_file}', index=False, engine='openpyxl')
print(f"✅ Updated file saved as: {output_file}")
