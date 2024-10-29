package excelinfo

import (
	"fmt"
	"os"
)

// ToAlphaString converts an index number to Excel column letters (e.g., 0 -> "A", 1 -> "B", 26 -> "AA")
func ToAlphaString(index int) string {
	result := ""
	for index >= 0 {
		remainder := index % 26
		result = string(rune('A'+remainder)) + result
		index = index/26 - 1
	}
	return result
}

// WriteHeadersToFile writes the headers of the sheet to a text file
func WriteHeadersToFile(headers []string, outputFilePath string) error {
	outputFile, err := os.Create(outputFilePath)
	if err != nil {
		return fmt.Errorf("failed to create output file: %w", err)
	}
	defer outputFile.Close()

	for index, header := range headers {
		column := ToAlphaString(index)
		line := fmt.Sprintf("%s = %d, ColumnName = %s\n", column, index, header)
		_, err := outputFile.WriteString(line)
		if err != nil {
			return fmt.Errorf("failed to write to output file: %w", err)
		}
	}

	return nil
}
