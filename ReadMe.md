excel2csv
==========


 Export a Sheet from a .xlsx file to a .csv file.
 excel2csv.exe excelfilename csvfilename [options]

 Note: excelfilename, csvfilename and sheet: is required.

 Options
 -------

 sheet : Sheet1 - (required) set the sheet name to export.

 row : 0        - the row number to start from

 col : 0        - the columns number to start from

 fields : "name1,name2,name3" - if set the field names will be the first row of the csv file.

 select : "A,B,C" - if set exports only the select columns names.

 Example:
 --------

     c:\data\export2csv data.xlsx data.csv sheet:Sheet1 row:1 col:1

also see 'export.bat' file for an example.
