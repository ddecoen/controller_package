# Controller Package

Month-end reporting package containing balance sheet, income statement, statement of cash flows, and flux analysis.

## Purpose

This package provides month-over-month financial reporting with variance analysis. The flux analysis compares the balance sheet and income statement with Month-over-Month variances.

The files are pulled from NetSuite and combined into a comprehensive Excel package.

## Setup Instructions

### Option 1: Go Version (Recommended)

#### Initial Setup (One-time)

1. **Install Go** (if not already installed):
   - Download from https://golang.org/dl/
   - Follow installation instructions for your OS

2. **Clone the repository:**
   ```bash
   git clone https://github.com/ddecoen/controller_package.git
   cd controller_package
   ```

3. **Build the application:**
   ```bash
   go build
   ```

4. **Verify the setup:**
   ```bash
   ./controller_package  # On macOS/Linux
   controller_package.exe  # On Windows
   ```

#### Running the Go Version

**On macOS/Linux:**
```bash
cd /path/to/controller_package
./controller_package
```

**On Windows:**
```cmd
cd C:\path\to\controller_package
controller_package.exe
```

**Alternative - Run without building:**
```bash
go run main.go
```

### Option 2: Python Version

#### Initial Setup (One-time)

1. **Clone the repository:**
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

#### Running the Python Version

```bash
python csv_to_excel_converter.py
# or
python3 csv_to_excel_converter.py
```

## Usage Workflow

1. **Export CSV files from NetSuite**
2. **Place CSV files in the controller_package directory**
3. **Open Terminal/Command Prompt and navigate to the project directory**
4. **Run the converter:**
   - Go version: `./controller_package` (or `controller_package.exe` on Windows)
   - Python version: `python csv_to_excel_converter.py`
5. **Review the generated Excel file:** `coder_financial_package.xlsx`

## Requirements

### Go Version:
- Go 1.16 or higher
- Git (for cloning the repository)
- CSV files exported from NetSuite

### Python Version:
- Python 3.6 or higher
- pandas library
- openpyxl library
- Git (for cloning the repository)
- CSV files exported from NetSuite

## Troubleshooting

### Go Version Issues:

- **"go: command not found"**: Install Go from https://golang.org/dl/
- **Build errors**: Run `go mod tidy` to ensure dependencies are correct
- **Permission denied**: Make sure the executable has run permissions (`chmod +x controller_package` on macOS/Linux)

### Python Version Issues:

- **"No CSV files found"**: Ensure your CSV files are in the same directory as the script
- **"Module not found" errors**: Install required packages with `pip install pandas openpyxl`
- **"Permission denied" errors**: Try using `pip install --user pandas openpyxl`
- **"Command not found: python"**: Try using `python3` instead of `python`

### General Issues:

- **No CSV files found**: Place CSV files in the same directory as the executable/script
- **Data formatting issues**: Ensure your CSV files follow the expected column structure
- **Excel file errors**: Check that you have write permissions in the directory

### Getting Help:

1. **Check Go installation:**
   ```bash
   go version
   ```

2. **Check Python installation:**
   ```bash
   python --version
   # or
   python3 --version
   ```

3. **List files in directory:**
   ```bash
   ls -la  # macOS/Linux
   dir     # Windows
   ```

## Performance

The Go version is significantly faster than the Python version, especially for large CSV files. It also produces a single executable that doesn't require additional dependencies to be installed.

## Go Version (New!)

We've added a Go version of the converter for better performance!

### Go Setup:

1. **Install Go** from https://golang.org/dl/
2. **Build the application:**
   ```bash
   go build
   ```
3. **Run the converter:**
   ```bash
   ./controller_package  # macOS/Linux
   controller_package.exe  # Windows
   ```

### Benefits of Go Version:
- Faster processing
- Single executable (no dependencies)
- Better error handling
- Cross-platform compatibility

### Alternative - Run without building:
```bash
go run main.go
```
