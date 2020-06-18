package main

import (
	"fmt"
	"os"
	"path/filepath"
	"strconv"
	"strings"
)

// Export a Sheet from a .xlsx file to a .csv file.
// excel2csv.exe excelfilename csvfilename [options]
//
// Note: excelfilename, csvfilename and sheet: is required.
//
// Options
// -------
// sheet:Sheet1 - (required) set the sheet name to export.
// row:0        - the row number to start from
// col:0        - the columns number to start from
// fields:"name1,name2,name3" - if set the field names will be the first row of the csv file.
//
// Example:
// --------
// export2csv data.xlsx data.csv sheet:Sheet1 row:1 col:1
func main() {

	if len(os.Args) < 3 {
		help()
		return
	}
	//os.Chdir(os.Getenv("=C:"))
	dir, err := os.Getwd()
	if err != nil {
		panic(err)
	}
	fmt.Println("Working Path: " + dir)

	export := new(Export)
	export.FromFile, _ = filepath.Abs(os.Args[1])
	export.ToFile, _ = filepath.Abs(os.Args[2])
	export.SheetName = "Sheet1"
	fmt.Println("excel2csv v0.01")
	for _, arg := range os.Args[3:] {
		parse(export, arg)
	}
	fmt.Printf("Exporting data: \n from %v \n to %v \n", export.FromFile, export.ToFile)
	export.Export()
	fmt.Println("Done")
}

func parse(export *Export, cmd string) {
	kv := strings.Split(cmd, ":")
	if len(kv) < 1 {
		panic("Invalid Option: " + cmd)
	}
	option := strings.ToLower(kv[0])
	switch option {
	case "sheet":
		export.SheetName = kv[1]
	case "row":
		v, err := strconv.ParseInt(kv[1], 10, 32)
		if err != nil {
			panic("Invalid Option: " + err.Error())
		}
		export.StartRow = int(v)
	case "col":
		v, err := strconv.ParseInt(kv[1], 10, 32)
		if err != nil {
			panic("Invalid Option: " + err.Error())
		}
		export.StartCol = int(v)
	case "fields":
		export.FieldNames = strings.Split(kv[1], ",")
	case "select":
		export.ExcelColumns = strings.Split(kv[1], ",")
	}
}

func help() {
	help := `Export a Sheet from a .xlsx file to a .csv file.
    excel2csv.exe excelfilename csvfilename [options]

    Note: excelfilename, csvfilename and sheet: is required.

    Options
    -------
    sheet:Sheet1 - (required) set the sheet name to export.
    row:0        - the row number to start from
    col:0        - the columns number to start from
    fields:"name1,name2,name3" - if set the field names will be the first row of the csv file.
		select:"A,B,C" - if set exports only the select columns names.

    Example:
    --------
    export2csv data.xlsx data.csv sheet:Sheet1 row:1 col:1
    `
	fmt.Print(help)
}
