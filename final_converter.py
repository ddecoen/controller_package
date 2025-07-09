#!/usr/bin/env python3
"""
CSV to Excel Financial Package Converter

This script combines multiple CSV files into a single Excel workbook.
Each CSV file becomes a separate worksheet in the Excel file.

Usage:
    python final_converter.py
    
The script will look for CSV files in the current directory and combine them.
"""

import pandas as pd
import os
import glob
from pathlib import Path

def clean_sheet_name(filename):
    """Convert filename to a clean Excel sheet name."""
    # Remove file extension and clean up
    name = Path(filename).stem
    
    # Replace underscores with spaces and title case
    name = name.replace('_', ' ').title()
    
    # Remove invalid Excel sheet name characters
    invalid_chars = ['[', ']', '*', '?', '/', '\\', ':']
    for char in invalid_chars:
        name = name.replace(char, '')
    
    # Ensure max 31 characters (Excel limit)
    return name[:31]

def combine_csv_to_excel(csv_files=None, output_file='financial_package.xlsx'):
    """
    Combine CSV files into a single Excel workbook.
    
    Args:
        csv_files: List of CSV files to process. If None, finds all CSV files.
        output_file: Name of output Excel file.
    """
    
    # If no files specified, find all CSV files in current directory
    if csv_files is None:
        csv_files = glob.glob('*.csv')
    
    # Filter to existing files
    existing_files = [f for f in csv_files if os.path.exists(f)]
    
    if not existing_files:
        print("No CSV files found to process.")
        return False
    
    print(f"Found {len(existing_files)} CSV files to process:")
    for f in existing_files:
        print(f"  - {f}")
    print()
    
    # Create Excel workbook
    try:
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            
            for csv_file in existing_files:
                try:
                    print(f"Processing: {csv_file}")
                    
                    # Read CSV file
                    df = pd.read_csv(csv_file)
                    
                    # Generate clean sheet name
                    sheet_name = clean_sheet_name(csv_file)
                    
                    # Write to Excel
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    
                    print(f"  ✓ Added as sheet: '{sheet_name}' ({len(df)} rows, {len(df.columns)} columns)")
                    
                except Exception as e:
                    print(f"  ✗ Error processing {csv_file}: {str(e)}")
                    continue
        
        print(f"\n✓ Successfully created: {output_file}")
        
        # Show file size
        if os.path.exists(output_file):
            file_size = os.path.getsize(output_file)
            print(f"File size: {file_size:,} bytes")
        
        return True
        
    except Exception as e:
        print(f"✗ Error creating Excel file: {str(e)}")
        return False

def main():
    """Main function."""
    print("CSV to Excel Financial Package Converter")
    print("=" * 50)
    
    # You can specify specific files here, or let it auto-detect
    # csv_files = ['income_statement.csv', 'balance_sheet.csv']
    csv_files = None  # Auto-detect all CSV files
    
    output_file = 'coder_financial_package.xlsx'
    
    success = combine_csv_to_excel(csv_files, output_file)
    
    if success:
        print("\n🎉 Conversion completed successfully!")
        print(f"\nYour financial package is ready: {output_file}")
        print("\nYou can now open this file in Excel to view:")
        print("  - Income Statement (with month-over-month variances)")
        print("  - Balance Sheet (with comparative analysis)")
        print("  - Any other financial data from your CSV files")
    else:
        print("\n❌ Conversion failed. Please check your CSV files.")

if __name__ == "__main__":
    main()
