"""
Export All Excel Sheets to Separate CSV Files
----------------------------------------------
This script reads an Excel workbook and exports each sheet
as a separate CSV file in the same directory.

Requirements:
    pip install pandas openpyxl

Usage:
    1. Update the 'excel_file' variable with your file path
    2. Run: python export_sheets_to_csv.py
"""

import pandas as pd
import os
import re

def sanitize_filename(name):
    """Remove or replace invalid characters for filenames."""
    # Replace invalid characters with underscores
    invalid_chars = r'[<>:"/\\|?*]'
    sanitized = re.sub(invalid_chars, '_', name)
    # Remove leading/trailing spaces and dots
    sanitized = sanitized.strip('. ')
    return sanitized

def export_sheets_to_csv(excel_file, output_folder=None):
    """
    Export all sheets from an Excel file to separate CSV files.
    
    Parameters:
        excel_file (str): Path to the Excel file
        output_folder (str): Optional output folder path. 
                            If None, uses the Excel file's directory.
    """
    # Validate file exists
    if not os.path.exists(excel_file):
        print(f"❌ Error: File not found: {excel_file}")
        return
    
    # Set output folder
    if output_folder is None:
        output_folder = os.path.dirname(excel_file)
        if output_folder == '':
            output_folder = '.'
    
    # Create output folder if it doesn't exist
    os.makedirs(output_folder, exist_ok=True)
    
    print(f"📂 Reading Excel file: {excel_file}")
    print(f"📁 Output folder: {output_folder}")
    print("-" * 50)
    
    try:
        # Read all sheets from the Excel file
        excel_data = pd.ExcelFile(excel_file, engine='openpyxl')
        sheet_names = excel_data.sheet_names
        
        print(f"📊 Found {len(sheet_names)} sheets to export:\n")
        
        exported_count = 0
        
        for sheet_name in sheet_names:
            try:
                # Read the sheet
                df = pd.read_excel(excel_data, sheet_name=sheet_name, header=None)
                
                # Create safe filename
                safe_name = sanitize_filename(sheet_name)
                csv_filename = f"{safe_name}.csv"
                csv_path = os.path.join(output_folder, csv_filename)
                
                # Export to CSV
                df.to_csv(csv_path, index=False, header=False, encoding='utf-8-sig')
                
                print(f"  ✅ {sheet_name}")
                print(f"     → {csv_filename} ({len(df)} rows)")
                
                exported_count += 1
                
            except Exception as e:
                print(f"  ❌ {sheet_name}: Error - {str(e)}")
        
        print("-" * 50)
        print(f"\n🎉 Successfully exported {exported_count} of {len(sheet_names)} sheets!")
        print(f"📁 CSV files saved to: {os.path.abspath(output_folder)}")
        
    except Exception as e:
        print(f"❌ Error reading Excel file: {str(e)}")

# ============================================================
# CONFIGURATION - UPDATE THIS PATH TO YOUR FILE
# ============================================================

if __name__ == "__main__":
    # Option 1: Specify the full path to your Excel file
    excel_file = r"c:\Users\susan\documents\Python Work\FY22_PLSFINAL.xlsx"
    
    # Option 2: If the script is in the same folder as the Excel file
    # excel_file = "NC_Public_Libraries.xlsx"
    
    # Optional: Specify a custom output folder for CSV files
    # output_folder = r"C:\Users\YourName\Documents\CSV_Exports"
    output_folder = None  # Uses same folder as Excel file
    
    # Run the export
    export_sheets_to_csv(excel_file, output_folder)