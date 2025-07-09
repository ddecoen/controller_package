#!/usr/bin/env python3
"""
Script to combine CSV files into a single Excel workbook.
Each CSV file becomes a separate worksheet.
"""

import pandas as pd
import os
import sys
from pathlib import Path

def combine_csv_to_excel(csv_files, output_file):
    """
    Combine multiple CSV files into a single Excel workbook.
    
    Args:
        csv_files: List of CSV file paths
        output_file: Output Excel file path
    """
    
    # Create Excel writer object
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        
        for csv_file in csv_files:
            if not os.path.exists(csv_file):
                print(f"Warning: File {csv_file} not found, skipping...")
                continue
                
            # Read CSV file
            try:
                df = pd.read_csv(csv_file)
                
                # Get sheet name from filename (without extension)
                sheet_name = Path(csv_file).stem
                
                # Ensure sheet name is valid for Excel (max 31 chars, no special chars)
                sheet_name = sheet_name[:31].replace('[', '').replace(']', '').replace('*', '').replace('?', '').replace('/', '').replace('\\', '')
                
                # Write to Excel sheet
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                
                print(f"Added {csv_file} as sheet '{sheet_name}'")
                
            except Exception as e:
                print(f"Error processing {csv_file}: {str(e)}")
                continue
    
    print(f"\nExcel file created: {output_file}")

def main():
    # Define input CSV files
    csv_files = [
        'income_statement.csv',
        'balance_sheet.csv'
    ]
    
    # Output Excel file
    output_file = 'coder_financial_package.xlsx'
    
    # Check if files exist
    existing_files = [f for f in csv_files if os.path.exists(f)]
    
    if not existing_files:
        print("No CSV files found. Please ensure the following files exist:")
        for f in csv_files:
            print(f"  - {f}")
        return
    
    print(f"Found {len(existing_files)} CSV files to process")
    
    # Combine CSV files into Excel
    combine_csv_to_excel(existing_files, output_file)

if __name__ == "__main__":
    main()
