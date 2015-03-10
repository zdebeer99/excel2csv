@echo off
echo Export someexcel.xlsx to somecsv.csv
excel2csv someexcel.xlsx somecsv.csv sheet:Sheet1 row:4 col:1 fields:"Col1,Col2,Col3,Col4"