package main

import (
	"encoding/csv"
	"fmt"
	"github.com/tealeg/xlsx"
	"os"
)

type Export struct {
	FromFile   string
	ToFile     string
	StartRow   int
	StartCol   int
	SheetName  string
	FieldNames []string
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
	if len(export.FieldNames) > 0 {
		table = append(table, export.FieldNames)
	}
	for rowi, row := range sheet.Rows {
		if rowi < export.StartRow {
			continue
		}

		if len(row.Cells) > export.StartCol {
			tablerow := make([]string, 30)
			table = append(table, tablerow)
			for coli, cell := range row.Cells[export.StartCol:] {
				tablerow[coli] = cell.Value
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
