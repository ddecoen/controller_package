package main

import (
	"encoding/csv"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strings"

	"github.com/xuri/excelize/v2"
)

func main() {
	fmt.Println("CSV to Excel Financial Package Converter")
	fmt.Println("Month-end Controller Package - Flux Analysis")
	fmt.Println()
	
	files, err := filepath.Glob("*.csv")
	if err != nil || len(files) == 0 {
		fmt.Println("No CSV files found")
		return
	}

	fmt.Printf("Found %d CSV files:\n", len(files))
	for _, file := range files {
		fmt.Printf("  - %s\n", file)
	}
	fmt.Println()

	f := excelize.NewFile()
	f.DeleteSheet("Sheet1")

	for _, csvFile := range files {
		fmt.Printf("Processing: %s\n", csvFile)
		
		file, err := os.Open(csvFile)
		if err != nil {
			fmt.Printf("Error opening %s: %v\n", csvFile, err)
			continue
		}
		
		reader := csv.NewReader(file)
		reader.FieldsPerRecord = -1
		
		var records [][]string
		for {
			record, err := reader.Read()
			if err == io.EOF {
				break
			}
			if err != nil {
				fmt.Printf("Error reading %s: %v\n", csvFile, err)
				break
			}
			records = append(records, record)
		}
		file.Close()

		sheetName := strings.TrimSuffix(csvFile, ".csv")
		sheetName = strings.ReplaceAll(sheetName, "_", " ")
		sheetName = strings.Title(sheetName)
		if len(sheetName) > 31 {
			sheetName = sheetName[:31]
		}

		index, err := f.NewSheet(sheetName)
		if err != nil {
			fmt.Printf("Error creating sheet: %v\n", err)
			continue
		}

		for rowIndex, row := range records {
			for colIndex, cell := range row {
				cellName, err := excelize.CoordinatesToCellName(colIndex+1, rowIndex+1)
				if err != nil {
					continue
				}
				f.SetCellValue(sheetName, cellName, cell)
			}
		}

		f.SetActiveSheet(index)
		fmt.Printf("Added %s as %s (%d rows)\n", csvFile, sheetName, len(records))
	}

	outputFile := "coder_financial_package.xlsx"
	if err := f.SaveAs(outputFile); err != nil {
		fmt.Printf("Error saving file: %v\n", err)
		return
	}

	fmt.Printf("\nSuccessfully created: %s\n", outputFile)
	fmt.Println("\nYour financial package contains:")
	fmt.Println("  - Income Statement (month-over-month variances)")
	fmt.Println("  - Balance Sheet (comparative analysis)")
	fmt.Println("  - Statement of Cash Flows (if provided)")
	fmt.Println("  - Flux analysis data")
}
