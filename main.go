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
	fmt.Println("CSV to Excel Converter - MoM Analysis")
	expectedFiles := getExpectedFiles()
	var availableFiles []FileMapping
	for _, fileMap := range expectedFiles {
		if checkFileExists(fileMap.Filename) {
			availableFiles = append(availableFiles, fileMap)
			fmt.Printf("Found: %s -> %s\n", fileMap.Filename, fileMap.SheetName)
		}
	}
	if len(availableFiles) == 0 {
		fmt.Println("No MoM_BS.csv or MoM_IS.csv files found")
		return
	}
	f := excelize.NewFile()
	f.DeleteSheet("Sheet1")
	for _, fileMap := range availableFiles {
		data, _ := readCSVFile(fileMap.Filename)
		index, _ := f.NewSheet(fileMap.SheetName)
		for rowIndex, row := range data {
			for colIndex, cell := range row {
				cellName, _ := excelize.CoordinatesToCellName(colIndex+1, rowIndex+1)
				f.SetCellValue(fileMap.SheetName, cellName, cell)
			}
		}
		f.SetActiveSheet(index)
		fmt.Printf("Added %s sheet\n", fileMap.SheetName)
	}
	f.SaveAs("coder_financial_package.xlsx")
	fmt.Println("Created coder_financial_package.xlsx")
}