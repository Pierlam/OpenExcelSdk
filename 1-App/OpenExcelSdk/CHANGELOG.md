## 0.8.0 Release (2026-XX-XX)

-todo: 

-45 unit tests, all are green.


## 0.7.1 Release (2026-03-28)

-Bugs fixed:  managed excelFile.WorkbookPart.WorkbookStylesPart is null.

-45 unit tests, all are green.


## 0.7.0 Release (2026-03-28)

-Improve scanning of rows and cells.

-Add GetLastRowAddress(): Get the last row address.

-Add GetRowAddress(): Get the row at address, not the index.

-Add GetRowCellsAtAddress(): Get the last row address, not the index

-Add GetLastColAddress(): Get the last cell col address in the row

-Rename GetRowAt to GetRowAtIndex

-Bugs fixing

-43 unit tests, all are green.


## 0.6.0 Release (2026-03-12)

-Add ExcelProcessor.CopyCellValue(); Copy a cell value to another one in another excel file.

-Add ExcelProcessor.GetRowCellsCount(sheet, rowIndex)

-Add GetRowAt(): Modify parameter rowIndex; now start from 1.

-Add SetCellValueCurrency() 

-Bugs fixing

-41 unit tests, all are green.


## 0.5.0 Release (2026-01-24)

-Add export of all styles of an Excel: ExcelProcessor.ExportAllStyles()

-Add GetRowCells(excelSheet, excelRow)

-Add GetRowCells(excelSheet, rowIndex)

-Get currency symbol from cell, when call GetCellValue method.

-Add GetCellColor(cell)  Get the cell color.

-Add SetCellColor(cell)  Set a color to a cell.

-Creation of the console application: OpenExcelExport.exe. 
  published here : https://pierlam.github.io/OpenExcelExport/

-Update OpenXML SDK to the lastest available version: 3.4.1

-24 unit tests, all are green.

## 0.4.0 Release (2025-12-30)

-Code Reworked, Simplification of several classes and methods.

-Many ExcelProcessor methods can now use friendly cell address like: A2.

-ExcelError class removed, replaced by a standard Exception class named ExcelException.

-Exception messages for each error code defined.

-Several Bugs fixed (number format,...).

-18 unit tests, all are green.