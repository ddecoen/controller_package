# Controller Package

Month-end reporting package containing balance sheet, income statement, statement of cash flows, and flux analysis.

## Purpose

This package provides month-over-month financial reporting with variance analysis. The flux analysis compares the balance sheet and income statement with Month-over-Month variances.

The files are pulled from NetSuite and combined into a comprehensive Excel package.

## Usage

1. Export CSV files from NetSuite
2. Place CSV files in this directory
3. Run: python csv_to_excel_converter.py
4. Review the generated Excel file

## Requirements

- Python 3.6+
- pandas
- openpyxl
