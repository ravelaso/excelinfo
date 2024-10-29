package main

import (
	"fmt"
	"log"
	"os"
	"path/filepath"
	"strings"

	"github.com/ravelaso/excelinfo"
	"github.com/xuri/excelize/v2"
)

func main() {
	if len(os.Args) < 2 {
		log.Fatalf("Usage: %s <input.xlsx>", os.Args[0])
	}
	// Get the input file name from the command-line arguments
	inputFile := os.Args[1]

	// Open the Excel file
	f, err := excelize.OpenFile(inputFile)
	if err != nil {
		log.Fatalf("Failed to open the Excel file: %v", err)
	}
	defer f.Close()

	// Get the list of all sheets
	sheets := f.GetSheetList()

	// Create a directory based on the input file name (without extension)
	baseName := strings.TrimSuffix(filepath.Base(inputFile), filepath.Ext(inputFile))
	outputDir := baseName
	err = os.Mkdir(outputDir, os.ModePerm)
	if err != nil && !os.IsExist(err) {
		log.Fatalf("Failed to create directory: %v", err)
	}

	// Loop through each sheet and process it
	for _, sheetName := range sheets {
		// Get all the rows in the current sheet
		rows, err := f.GetRows(sheetName)
		if err != nil {
			log.Printf("Failed to get rows for sheet %s: %v", sheetName, err)
			continue
		}

		if len(rows) == 0 {
			log.Printf("No rows found in sheet %s", sheetName)
			continue
		}

		// Assuming the first row contains the headers
		headers := rows[0]

		// Create the output file path
		outputFilePath := filepath.Join(outputDir, fmt.Sprintf("%s.txt", sheetName))

		// Write the headers to the output file
		err = excelinfo.WriteHeadersToFile(headers, outputFilePath)
		if err != nil {
			log.Printf("Failed to write headers for sheet %s: %v", sheetName, err)
			continue
		}

		fmt.Printf("Output file generated successfully for sheet: %s\n", sheetName)
	}
}
