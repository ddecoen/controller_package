# Controller Package

Month-end reporting package containing balance sheet, income statement, statement of cash flows, and flux analysis.

## Purpose

This package provides month-over-month financial reporting with variance analysis. The flux analysis compares the balance sheet and income statement with Month-over-Month variances.

The files are pulled from NetSuite and combined into a comprehensive Excel package.

## Setup Instructions

### Initial Setup (One-time)

1. **Clone the repository to your local machine:**
   ```bash
   git clone https://github.com/ddecoen/controller_package.git
   cd controller_package
   ```

2. **Install Python dependencies:**
   ```bash
   pip install pandas openpyxl
   ```
   
   *If you encounter permission issues, try:*
   ```bash
   pip install --user pandas openpyxl
   ```

3. **Verify the setup:**
   ```bash
   python csv_to_excel_converter.py
   ```
   *This should display a message about no CSV files found (which is expected initially)*

### Accessing via Terminal

**On macOS/Linux:**
```bash
# Navigate to the project directory
cd /path/to/controller_package

# Run the converter
python csv_to_excel_converter.py
```

**On Windows:**
```cmd
# Navigate to the project directory
cd C:\path\to\controller_package

# Run the converter
python csv_to_excel_converter.py
```

**Alternative using Python 3 explicitly:**
```bash
python3 csv_to_excel_converter.py
```

## Usage Workflow

1. **Export CSV files from NetSuite**
2. **Place CSV files in the controller_package directory**
3. **Open Terminal/Command Prompt and navigate to the project directory**
4. **Run the converter:**
   ```bash
   python csv_to_excel_converter.py
   ```
5. **Review the generated Excel file:** `coder_financial_package.xlsx`

## Requirements

- Python 3.6 or higher
- pandas library
- openpyxl library
- Git (for cloning the repository)
- CSV files exported from NetSuite

## Troubleshooting

### Common Issues:

- **"No CSV files found"**: Ensure your CSV files are in the same directory as the script
- **"Module not found" errors**: Install required packages with `pip install pandas openpyxl`
- **"Permission denied" errors**: Try using `pip install --user pandas openpyxl`
- **"Command not found: python"**: Try using `python3` instead of `python`

### Getting Help:

1. **Check Python installation:**
   ```bash
   python --version
   # or
   python3 --version
   ```

2. **Check if packages are installed:**
   ```bash
   python -c "import pandas, openpyxl; print('All packages installed successfully')"
   ```

3. **List files in directory:**
   ```bash
   ls -la  # macOS/Linux
   dir     # Windows
   ```
