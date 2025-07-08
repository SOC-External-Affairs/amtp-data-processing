import os
import pandas as pd

# === CONFIGURATION ===
input_file = './inbox/data.xlsx'                     # Your original Excel file
output_file = 'updated_exported_data.xlsx'  # Output file with matches
search_root = './inbox/'                             # Folder to search in
# === LOAD EXCEL FILE ===
try:
    df = pd.read_excel(input_file, engine='openpyxl')
except FileNotFoundError:
    raise FileNotFoundError(f"Could not find the file: {input_file}")

# === CHECK FOR RESPONSE ID COLUMN ===
response_id_col = None
for col in df.columns:
    if 'response' in col.lower() or 'id' in col.lower():
        response_id_col = col
        break

if not response_id_col:
    print("Available columns:", list(df.columns))
    raise ValueError("No Response ID column found.")

# === FUNCTION TO FIND MATCHING FILES ===
def find_matching_files(response_id, root_dir):
    matches = []
    for dirpath, _, filenames in os.walk(root_dir):
        for filename in filenames:
            if str(response_id) in filename:
                full_path = os.path.join(dirpath, filename)
                matches.append(full_path)
    return matches

# === APPLY TO EACH ROW ===
df['Matched Files'] = df[response_id_col].apply(
    lambda rid: '| '.join(find_matching_files(rid, search_root))
)

# === SAVE UPDATED FILE ===
os.makedirs('outbox', exist_ok=True)
df.to_excel(f'outbox/{output_file}', index=False, engine='openpyxl')
print(f"✅ Updated file saved as: {output_file}")
