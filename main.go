package main

import (
	"encoding/csv"
	"fmt"
	"io"
	"os"

	"github.com/xuri/excelize/v2"
)

type FileMapping struct {
	Filename  string
	SheetName string
}

func getExpectedFiles() []FileMapping {
	return []FileMapping{
		{"MoM_BS.csv", "Balance Sheet"},
		{"MoM_IS.csv", "Income Statement"},
	}
}

func checkFileExists(filename string) bool {
	_, err := os.Stat(filename)
	return !os.IsNotExist(err)
}

func readCSVFile(filename string) ([][]string, error) {
	file, err := os.Open(filename)
	if err != nil {
		return nil, err
	}
	defer file.Close()

	reader := csv.NewReader(file)
	reader.FieldsPerRecord = -1

	// Skip the first 10 rows (rows 1-10), start reading from row 11
	for i := 0; i < 10; i++ {
		_, err := reader.Read()
		if err == io.EOF {
			// If file has fewer than 10 rows, return empty result
			return [][]string{}, nil
		}
		if err != nil {
			return nil, err
		}
	}

	// Now read the actual data starting from row 11
	var records [][]string
	for {
		record, err := reader.Read()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, err
		}
		records = append(records, record)
	}
	return records, nil
}

func main() {
	fmt.Println("CSV to Excel Financial Package Converter")
	fmt.Println("Month-end Controller Package - Flux Analysis")
	fmt.Println("Processing MoM_BS.csv and MoM_IS.csv (skipping first 10 rows)")
	fmt.Println()
	
	expectedFiles := getExpectedFiles()
	var availableFiles []FileMapping
	for _, fileMap := range expectedFiles {
		if checkFileExists(fileMap.Filename) {
			availableFiles = append(availableFiles, fileMap)
			fmt.Printf("✓ Found: %s -> %s\n", fileMap.Filename, fileMap.SheetName)
		} else {
			fmt.Printf("✗ Missing: %s\n", fileMap.Filename)
		}
	}
	if len(availableFiles) == 0 {
		fmt.Println("\nNo MoM_BS.csv or MoM_IS.csv files found.")
		fmt.Println("Please ensure these files are in the current directory.")
		fmt.Println("Note: Data will be read starting from row 11 (skipping first 10 rows).")
		return
	}
	f := excelize.NewFile()
	f.DeleteSheet("Sheet1")
	fmt.Println("\nProcessing files (skipping first 10 rows in each file):")
	for _, fileMap := range availableFiles {
		fmt.Printf("\nProcessing: %s\n", fileMap.Filename)
		
		data, err := readCSVFile(fileMap.Filename)
		if err != nil {
			fmt.Printf("  Error reading %s: %v\n", fileMap.Filename, err)
			continue
		}
		
		index, err := f.NewSheet(fileMap.SheetName)
		if err != nil {
			fmt.Printf("  Error creating sheet: %v\n", err)
			continue
		}
		
		for rowIndex, row := range data {
			for colIndex, cell := range row {
				cellName, _ := excelize.CoordinatesToCellName(colIndex+1, rowIndex+1)
				f.SetCellValue(fileMap.SheetName, cellName, cell)
			}
		}
		
		f.SetActiveSheet(index)
		fmt.Printf("  ✓ Added as '%s' sheet (%d data rows, starting from row 11)\n", fileMap.SheetName, len(data))
	}
	f.SaveAs("coder_financial_package.xlsx")
	fmt.Println("Created coder_financial_package.xlsx")
}