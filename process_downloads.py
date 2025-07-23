import os
import zipfile
import shutil
import glob
from datetime import datetime, timedelta
from pathlib import Path
from settings import DOWNLOADS_INBOX_PATH, DOWNLOADS_INBOX_DATA_PATH, DOWNLOADS_TIME_WINDOW_HOURS, DOWNLOADS_FILE_PATTERN, DOWNLOADS_TARGET_XLSX

def process_amtp_downloads():
    """
    Processes recent AMTP zip files from Downloads folder.
    
    Steps:
    1. Find AMTP zip files created in last 2 hours
    2. Extract their contents
    3. Move XLSX file to ./inbox/data.xlsx
    4. Move other files to ./inbox/data/
    """
    
    downloads_path = Path.home() / "Downloads"
    inbox_path = Path(DOWNLOADS_INBOX_PATH)
    inbox_data_path = Path(DOWNLOADS_INBOX_DATA_PATH)
    
    # Create directories if they don't exist
    inbox_path.mkdir(exist_ok=True)
    inbox_data_path.mkdir(exist_ok=True)
    
    # Debug: List all zip files in Downloads
    all_zips = list(downloads_path.glob("*.zip"))
    print(f"\n=== DEBUG: Scanning Downloads Directory ===")
    print(f"[INFO] Found {len(all_zips)} total zip files")
    for zip_file in all_zips: 
        print(f"[FILE] {zip_file.name}")
    print("===============================\n")    
    # Find AMTP zip files from configured time window
    cutoff_time = datetime.now() - timedelta(hours=DOWNLOADS_TIME_WINDOW_HOURS)
    amtp_zips = []
    
    # Look for files matching pattern
    for zip_file in downloads_path.glob(DOWNLOADS_FILE_PATTERN):
        file_time = datetime.fromtimestamp(zip_file.stat().st_mtime)
        print(f"📅 {zip_file.name}: {file_time} (cutoff: {cutoff_time})")
        if file_time > cutoff_time:
            amtp_zips.append((zip_file, file_time))
            print(f"  ✅ Added to processing list")
        else:
            print(f"  ❌ Too old")
    
    if not amtp_zips:
        print("📭 No recent AMTP zip files found in Downloads")
        print(f"💡 Try extending time window or check file names contain 'AMTP'")
        return
    
    # Sort by creation time, process all recent files
    amtp_zips.sort(key=lambda x: x[1], reverse=True)
    
    print(f"📦 Processing {len(amtp_zips)} AMTP zip files...")
    
    # Process each zip file
    for zip_file, file_time in amtp_zips:
        print(f"\n📦 Processing: {zip_file.name}")
        process_single_zip(zip_file, inbox_path, inbox_data_path)

def process_single_zip(zip_file, inbox_path, inbox_data_path):
    
    """
    Process a single zip file - extract and organize contents.
    """
    temp_extract_path = inbox_path / f"temp_extract_{zip_file.stem}"
    temp_extract_path.mkdir(exist_ok=True)
    
    try:
        with zipfile.ZipFile(zip_file, 'r') as zip_ref:
            zip_ref.extractall(temp_extract_path)
        
        print(f"📂 Extracted to temporary folder")
        
        # Debug: List extracted files
        extracted_files = list(temp_extract_path.rglob("*"))
        print(f"📁 Extracted {len(extracted_files)} items:")
        for item in extracted_files[:10]:  # Show first 10
            print(f"  - {item.relative_to(temp_extract_path)} ({'file' if item.is_file() else 'dir'})")
        
        # Find and move XLSX file (if any)
        xlsx_files = list(temp_extract_path.rglob("*.xlsx"))
        if xlsx_files:
            xlsx_file = xlsx_files[0]  # Take first XLSX found
            target_xlsx = inbox_path / DOWNLOADS_TARGET_XLSX
            
            # Remove existing data.xlsx if it exists
            if target_xlsx.exists():
                target_xlsx.unlink()
            
            shutil.move(str(xlsx_file), str(target_xlsx))
            print(f"📊 Moved XLSX to: {target_xlsx}")
        
        # Move all other files (including non-XLSX files) to ./inbox/data/
        moved_count = 0
        for item in temp_extract_path.rglob("*"):
            if item.is_file() and item.suffix.lower() != ".xlsx":
                # Create subdirectory structure in data folder
                relative_path = item.relative_to(temp_extract_path)
                target_path = inbox_data_path / relative_path
                
                print(f"📦 Moving: {relative_path} -> {target_path}")
                
                # Create parent directories if needed
                target_path.parent.mkdir(parents=True, exist_ok=True)
                
                # Move file
                shutil.move(str(item), str(target_path))
                moved_count += 1
        
        print(f"📁 Moved {moved_count} files to: {inbox_data_path}")
        
        # Clean up temp directory
        shutil.rmtree(temp_extract_path)
        print("🧹 Cleaned up temporary files")
        
    except zipfile.BadZipFile:
        print(f"❌ Error: {zip_file.name} is not a valid zip file")
    except Exception as e:
        print(f"❌ Error processing zip file: {e}")
        # Clean up temp directory on error
        if temp_extract_path.exists():
            shutil.rmtree(temp_extract_path)

if __name__ == "__main__":
    print("🔍 Looking for recent AMTP zip files...")
    process_amtp_downloads()
    print("✅ Processing complete")