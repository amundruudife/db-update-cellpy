#!/usr/bin/env python3
"""
Super Simple SharePoint File Copy - Log Sheet Only

Finds newest Cell_Log.xlsx in Downloads, extracts only the "log" sheet,
and saves it to source_data/Cell_Log.xlsx

Usage: python copy_sharepoint_file.py
"""

import shutil
import pandas as pd
from pathlib import Path
from datetime import datetime
from .common_utils import ensure_directory

def copy_cell_log_to_source_data():
    """
    Find newest Cell_Log.xlsx in Downloads, extract "log" sheet,
    and save to source_data/Cell_Log.xlsx
    
    Returns:
        bool: True if successful, False if failed
    """
    
    print("🔍 Looking for Cell_Log.xlsx in Downloads...")
    
    # Candidate locations: synced SharePoint folder first, then Downloads
    sharepoint_folder = Path.home() / "Institutt for Energiteknikk" / "Users of IFE Battery Lab - 03_BaDeLab"
    downloads_folder = Path.home() / "Downloads"
    cell_log_files = []

    for folder in [sharepoint_folder, downloads_folder]:
        if folder.exists():
            matches = list(folder.glob("Cell_Log*.xlsx"))
            cell_log_files.extend(matches)
        else:
            print(f"ℹ️  Skipping missing folder: {folder}")
    
    if not cell_log_files:
        print("❌ No Cell_Log*.xlsx files found in SharePoint sync folder or Downloads")
        print(f"   Checked: {sharepoint_folder}")
        print(f"   Checked: {downloads_folder}")
        print("   💡 Ensure the file is synced or downloaded locally")
        return False
    
    # Find the newest file
    newest_file = max(cell_log_files, key=lambda f: f.stat().st_mtime)
    
    # Show what we found
    file_age_hours = (datetime.now().timestamp() - newest_file.stat().st_mtime) / 3600
    file_size = newest_file.stat().st_size
    
    print(f"✅ Found: {newest_file.name}")
    print(f"   Size: {file_size} bytes")
    print(f"   Age: {file_age_hours:.1f} hours")
    if file_age_hours > 24:
        print("⚠️  Warning: Downloads file is older than 24 hours - may be stale. Download a fresh copy from SharePoint.")
    
    # Destination file
    destination = Path("source_data/Cell_Log.xlsx")
    
    try:
        # Create source_data directory if needed
        ensure_directory(destination.parent)
        
        print(f"\n📋 Extracting 'log' sheet from: {newest_file.name}")
        
        # Read only the "log" sheet
        try:
            log_df = pd.read_excel(newest_file, sheet_name='log')
            print(f"   ✅ Read 'log' sheet: {len(log_df)} rows, {len(log_df.columns)} columns")
            # Compute max key to surface stale sources quickly
            max_key = pd.to_numeric(log_df.iloc[:, 0], errors='coerce').max()
            print(f"   🔑 Max key detected: {max_key}")
            
            # Debug: Check for file_name_indicator column
            cols_with_file = [col for col in log_df.columns if 'file' in str(col).lower()]
            if cols_with_file:
                print(f"   📋 Columns containing 'file': {cols_with_file}")
            else:
                print(f"   ⚠️  No columns containing 'file' found")
                
            # Debug: Show first few column names
            print(f"   📋 First 10 columns: {list(log_df.columns[:10])}")
            
        except ValueError as e:
            if "Worksheet named 'log'" in str(e):
                print("   ❌ No 'log' sheet found in Excel file")
                print("   💡 Check that the file has a sheet named 'log'")
                return False
            else:
                print(f"   ❌ Error reading Excel file: {e}")
                return False
        except Exception as e:
            print(f"   ❌ Error reading Excel file: {e}")
            return False
        
        # Save only the log sheet to destination
        print(f"   📝 Saving to: {destination}")
        log_df.to_excel(destination, sheet_name='log', index=False)
        
        # Verify the saved file
        if destination.exists():
            dest_size = destination.stat().st_size
            print(f"✅ Copy successful: {dest_size} bytes")
            print(f"   From: {newest_file} (log sheet only)")
            print(f"   To: {destination.absolute()}")
            
            # Quick validation of saved file
            try:
                verify_df = pd.read_excel(destination, sheet_name='log')
                if len(verify_df) == len(log_df):
                    print("✅ Verification passed - log sheet extracted correctly")
                    return True
                else:
                    print("⚠️  Row count mismatch in saved file")
                    return False
            except Exception as e:
                print(f"⚠️  Could not verify saved file: {e}")
                return True  # File was saved, verification failed but might still be OK
        else:
            print("❌ Copy failed - destination file not created")
            return False
            
    except Exception as e:
        print(f"❌ Copy failed: {e}")
        return False

def main():
    """Main function"""
    
    print("Simple SharePoint File Copy - Log Sheet Only")
    print("=" * 50)
    print("This extracts the 'log' sheet from the newest Cell_Log.xlsx")
    print("in Downloads and saves it to source_data/Cell_Log.xlsx")
    print()
    
    success = copy_cell_log_to_source_data()
    
    if success:
        print("\n🎉 SUCCESS!")
        print("   'log' sheet extracted and ready in source_data/Cell_Log.xlsx")
        print("   You can now run your database update script")
    else:
        print("\n❌ FAILED!")
        print("   Steps to fix:")
        print("   1. Go to SharePoint in your browser")
        print("   2. Download Cell_Log.xlsx (it goes to Downloads folder)")
        print("   3. Run this script again")

if __name__ == "__main__":
    main() 