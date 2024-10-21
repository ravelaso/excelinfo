package main

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"log"
	"os"
)

// ToAlphaString converts an index number to Excel column letters (e.g., 0 -> "A", 1 -> "B", 26 -> "AA")
func ToAlphaString(index int) string {
	result := ""
	for index >= 0 {
		remainder := index % 26
		result = string('A'+remainder) + result
		index = index/26 - 1
	}
	return result
}

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

	// Get all the rows in the first sheet
	rows, err := f.GetRows(f.GetSheetName(0))
	if err != nil {
		log.Fatalf("Failed to get rows: %v", err)
	}

	if len(rows) == 0 {
		log.Fatalf("No rows found in the sheet")
	}

	// Assuming the first row contains the headers
	headers := rows[0]

	// Create an output file
	outputFile, err := os.Create("output.txt")
	if err != nil {
		log.Fatalf("Failed to create output file: %v", err)
	}
	defer outputFile.Close()

	// Write the headers with their index
	for index, header := range headers {
		column := ToAlphaString(index)
		line := fmt.Sprintf("%s = %d, ColumnName = %s\n", column, index, header)
		_, err := outputFile.WriteString(line)
		if err != nil {
			log.Fatalf("Failed to write to output file: %v", err)
		}
	}

	fmt.Println("Output file generated successfully.")
}
