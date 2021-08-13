package main

import (
	"encoding/csv"
	"fmt"
	"log"
	"os"

	"github.com/tealeg/xlsx"
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
	if sheet == nil {
		log.Fatalf("Could not find sheet '%s' in excel workbook.", export.SheetName)
		return
	}

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
		numberOfCols = len(sheet.Cols)
	}

	fmt.Println("Columns Found : ", numberOfCols)
	fmt.Println("Excel Columns : ", len(sheet.Cols))

	if len(export.FieldNames) != len(export.ExcelColumns) {
		log.Fatalln("Field names and excel columns must contain the same number of fields.")
		return
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
				cellIndex := coli
				if len(export.ExcelColumns) > 0 {
					cellIndex = xlsx.ColLettersToIndex(export.ExcelColumns[coli])
				}
				if cellIndex >= len(row.Cells) {
					fmt.Printf("Row %v Column %v ignored, out of index rage %v, max columns %v \r\n", rowi+1, export.ExcelColumns[coli], cellIndex, len(row.Cells))
					continue
				}
				cell := row.Cells[cellIndex]
				if cell.IsTime() {
					t1, err := cell.GetTime(false)
					if err != nil {
						tablerow[coli] = fmt.Sprint("%s", err)
					} else {
						tablerow[coli] = t1.Format("2006-01-02 15:04")
					}
				} else {
					tablerow[coli] = cell.Value
				}
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
