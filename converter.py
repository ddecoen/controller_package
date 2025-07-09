import pandas as pd
import sys
import os

def combine_csv_to_excel(csv_files, output_file):
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        for csv_file in csv_files:
            if os.path.exists(csv_file):
                df = pd.read_csv(csv_file)
                sheet_name = os.path.splitext(csv_file)[0].replace("_", " ").title()[:31]
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                print(f"Added {csv_file} as {sheet_name}")
    print(f"Created {output_file}")

if __name__ == "__main__":
    csv_files = ["income_statement.csv", "balance_sheet.csv"]
    combine_csv_to_excel(csv_files, "financial_package.xlsx")
