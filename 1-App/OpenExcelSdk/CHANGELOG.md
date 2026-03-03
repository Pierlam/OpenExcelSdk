## 0.6.0 Release (2026-03-XX)

-Add ExcelProcessor.CopyCellValue() from a cell to another one in another excel file.

-Add ExcelProcessor.GetRowCellsCount(sheet, rowIndex)

-GetRowAt(): Modify parameter rowIndex; now start from 1.


-30 unit tests, all are green.


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