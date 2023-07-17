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

func copyCellStyles(sourceFile, destFile *excelize.File, sourceSheet, destSheet, srcCell, destCell string) error {
	srcStyle, err := sourceFile.GetCellStyle(sourceSheet, srcCell)
	if err != nil {
		return err
	}

	// Create a new style in the destination workbook
	destStyle, _ := destFile.NewStyle()
	*destStyle = *srcStyle

	// Apply the style to the destination cell
	if err := destFile.SetCellStyle(destSheet, destCell, destCell, destStyle); err != nil {
		return err
	}

	return nil
}

func main() {
	args := os.Args[1:] // Skip the first argument as it contains the program name
	if len(args) < 2 {
		fmt.Println("Usage: go run main.go <source workbook> <destination workbook>")
		return
	}

	sourceWorkbook := args[0]
	destinationWorkbook := args[1]
	fmt.Println("Source workbook:", sourceWorkbook)
	fmt.Println("Destination workbook:", destinationWorkbook)

	// Open the source workbook
	sourceFile, err := excelize.OpenFile(sourceWorkbook)
	if err != nil {
		fmt.Println("Failed to open source workbook:", err)
		return
	}
	defer sourceFile.Close()

	// Open the destination workbook
	destFile := excelize.NewFile()

	// Copy sheets from the source workbook to the destination workbook
	sourceSheets := sourceFile.GetSheetMap()
	for _, sourceSheet := range sourceSheets {
		// Read data from the source workbook
		rows, err := sourceFile.GetRows(sourceSheet)
		if err != nil {
			fmt.Printf("Failed to get rows from source sheet '%s': %v\n", sourceSheet, err)
			continue
		}

		// Create a new sheet in the destination workbook
		destSheet := sourceSheet
		destFile.NewSheet(destSheet)

		// Copy content and styles from source sheet to destination sheet
		for rowIndex, row := range rows {
			for colIndex, cellValue := range row {
				// Set the cell value in the destination workbook
				colLetter := columnNumberToLetter(colIndex + 1)
				destCell := colLetter + fmt.Sprint(rowIndex+1)
				destFile.SetCellValue(destSheet, destCell, cellValue)

				// Copy cell style from source to destination
				srcCell := colLetter + fmt.Sprint(rowIndex+1)
				if err := copyCellStyles(sourceFile, destFile, sourceSheet, destSheet, srcCell, destCell); err != nil {
					fmt.Println("Failed to copy cell style:", err)
				}
			}
		}
	}

	// Save the destination workbook
	if err := destFile.SaveAs(destinationWorkbook); err != nil {
		fmt.Println("Failed to save destination workbook:", err)
		return
	}

	fmt.Println("Sheets copied and pasted successfully.")
}
