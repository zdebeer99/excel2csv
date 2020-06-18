package main

import (
	"encoding/csv"
	"fmt"
	"github.com/tealeg/xlsx"
	"os"
)

type Export struct {
	FromFile     string
	ToFile       string
	StartRow     int
	StartCol     int
	SheetName    string
	FieldNames   []string
	ExcelColumns []string
}

func NewExport(fromfile, tofile, sheetname string) *Export {
	return &Export{
		FromFile:  fromfile,
		ToFile:    tofile,
		SheetName: sheetname,
	}
}

func (export *Export) Export() {

	//Open xlsx file for reading
	xlFile, error := xlsx.OpenFile(export.FromFile)
	if error != nil {
		fmt.Println("File Error:", error)
		return
	}

	sheet := xlFile.Sheet[export.SheetName]
	var table [][]string
	var numberOfCols int
	numberOfCols = 0
	if len(export.FieldNames) > 0 {
		table = append(table, export.FieldNames)
		numberOfCols = len(export.FieldNames)
		if len(export.ExcelColumns) == 0 {
			export.ExcelColumns = make([]string, numberOfCols)
			for n := 0; n < numberOfCols; n++ {
				export.ExcelColumns[n] = xlsx.ColIndexToLetters(n + export.StartCol)
			}
		}
	} else if len(export.ExcelColumns) > 0 {
		numberOfCols = len(export.ExcelColumns)
		export.FieldNames = export.ExcelColumns
	} else {
		numberOfCols = 0
	}
	if len(export.FieldNames) != len(export.ExcelColumns) {
		panic("Field names and excel columns must contain the same number of fields.")
	}
	for rowi, row := range sheet.Rows {
		if rowi < export.StartRow {
			continue
		}

		if numberOfCols == 0 {
			numberOfCols = len(row.Cells)
		}

		if len(row.Cells) > export.StartCol {
			tablerow := make([]string, numberOfCols)
			table = append(table, tablerow)
			// for coli, cell := range row.Cells[export.StartCol : export.StartCol+numberOfCols] {
			// 	tablerow[coli] = cell.Value
			// }
			for coli := 0; coli < numberOfCols; coli++ {
				tablerow[coli] = row.Cells[xlsx.ColLettersToIndex(export.ExcelColumns[coli])].Value
			}
		}
	}

	//Create csv file for export.
	csvfile, err := os.Create(export.ToFile)
	defer csvfile.Close()
	if err != nil {
		fmt.Println(err)
		return
	}

	writer := csv.NewWriter(csvfile)
	writer.WriteAll(table)
}
