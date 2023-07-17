package main

import (
	"fmt"
	"os"

	"github.com/xuri/excelize/v2"
)

func columnNumberToLetter(columnNumber int) string {
	columnLetter := ""
	for columnNumber > 0 {
		remainder := (columnNumber - 1) % 26
		columnLetter = string('A'+remainder) + columnLetter
		columnNumber = (columnNumber - 1) / 26
	}
	return columnLetter
}

func main() {
	args := os.Args[1:] // Skip the first argument as it contains the program name
	if len(args) <= 0 {
		fmt.Println("Usage: go run main.go <workbook>")
		return
	}

	workbook := args[0]
	fmt.Println(workbook)

	// Open the Excel file
	f, err := excelize.OpenFile(workbook)
	if err != nil {
		fmt.Println("Failed to open workbook:", err)
		return
	}
	defer func() {
		if err := f.SaveAs(workbook); err != nil {
			fmt.Println("Failed to save workbook:", err)
		}
	}()

	sourceSheet := "Settings"
	destinationSheet := "DestinationSheet"

	// Read data from the source sheet
	rows, err := f.GetRows(sourceSheet)
	if err != nil {
		fmt.Println("Failed to get rows from source sheet:", err)
		return
	}

	// Create a new destination sheet
	destinationIndex, _ := f.NewSheet(destinationSheet)
	if destinationIndex == 0 {
		fmt.Println("Failed to create destination sheet")
		return
	}

	// Activate the destination sheet using the index
	f.SetActiveSheet(destinationIndex)

	// Write data and style to the destination sheet
	for rowIndex, row := range rows {
		for colIndex, cellValue := range row {
			// Set the cell value in the destination sheet
			colLetter := columnNumberToLetter(colIndex + 1)
			destCell := colLetter + fmt.Sprint(rowIndex+1)
			f.SetCellValue(destinationSheet, destCell, cellValue)

			// Copy cell style from source to destination
			style, err := f.GetCellStyle(sourceSheet, colLetter+fmt.Sprint(rowIndex+1))
			if err != nil {
				fmt.Println("Failed to get cell style:", err)
			} else {
				err = f.SetCellStyle(destinationSheet, destCell, destCell, style)
				if err != nil {
					fmt.Println("Failed to set cell style:", err)
				}
			}
		}
	}

	fmt.Println("Data copied and pasted successfully.")
}
