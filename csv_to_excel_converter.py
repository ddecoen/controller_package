import pandas as pd
import os
import glob
from pathlib import Path

def combine_csv_to_excel():
    csv_files = glob.glob("*.csv")
    if not csv_files:
        print("No CSV files found")
        return
    
    with pd.ExcelWriter("coder_financial_package.xlsx", engine="openpyxl") as writer:
        for csv_file in csv_files:
            df = pd.read_csv(csv_file)
            sheet_name = Path(csv_file).stem.replace("_", " ").title()[:31]
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            print(f"Added {csv_file} as {sheet_name}")
    
    print("Created coder_financial_package.xlsx")

if __name__ == "__main__":
    combine_csv_to_excel()
