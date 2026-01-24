## 0.5.0 Release (2026-01-24)

-Add export of all styles of an Excel: ExcelProcessor.ExportAllStyles()

-Add new methods: GetRowCells(excelSheet, excelRow)

-Add new methods: GetRowCells(excelSheet, rowIndex)

-Get currency symbol from cell, when call GetCellValue method.

-GetCellColor(cell)  Get the background and foreground cell color.

-Creation of the console applcation: OpenExcelExport.exe. 
  published here : https://pierlam.github.io/OpenExcelExport/


## 0.4.0 Release (2025-12-30)

-Code Reworked, Simplification of several classes and methods.

-Many ExceProcessor methods can now use friendly cell address like: A2.

-ExcelError class removed, replaced by a standard Exception class named ExcelException.

-Exception messages for each error code defined.

-Several Bugs fixed (number format,...).

-24 unit tests, all are green.