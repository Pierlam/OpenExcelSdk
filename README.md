# What is OpenExcelSdk ?

OpenExcelSdk is a backend dotnet library to use Excel (xlsx) very easily.

It's an open-source .NET library written in C#/NET8.

The only dependency is OpenXML SDK, the official Microsoft library to work with Excel files.

This Microsoft library is not easy to use, so OpenExcelSdk propose a simple way to use Excel rows and cells values.

# A quick example

## Read B2 string cell value

The code below open an excel file, get the first sheet, then get the B2 cell.
After that, if the type of the cell is string, get the string cell value and display it.

```
bool res;
ExcelError error;
ExcelCellValueMulti cellValueMulti;

// create the excel processor to read/create/update cell
ExcelProcessor proc = new ExcelProcessor();

// open an excel file
string filename = PathFiles + "GetCellTypeAndValueCustom.xlsx";
res = proc.Open(filename, out ExcelFile excelFile, out error);

// get the first sheet of the excel file
proc.GetFirstSheet(excelFile, out ExcelSheet excelSheet, out error);


// get cell at B2 
proc.GetCellAt(excelSheet, 2, 2, out ExcelCell cell, out error);

// get the type and the value of cell
proc.GetCellTypeAndValue(excelSheet, cell, out cellValueMulti, out error);

// type: of the cell can be: string, int, double, DateTime, DateOnly, TimeOnly.

if(cellValueMulti.CellType== ExcelCellType.String)
{
  Console.WriteLine("B2 type: string, value:" + proc.GetCellValueAsString(excelFile, cell));
} 
```


## others functions

# Get row/last row index

The code below get a the last row index, and also get a row at an index.

```
ExcelError error;
bool res;
ExcelProcessor proc = new ExcelProcessor();

string filename = @".\Files\data.xlsx";
proc.Open(filename, out ExcelFile excelFile, out error);
proc.GetSheetAt(excelFile, 0, out ExcelSheet excelSheet, out error);

int lastRowIdx = proc.GetLastRowIndex(excelSheet);
Console.WriteLine("last row idx: " + lastRowIdx);

res = proc.GetRowAt(excelSheet, 0, out ExcelRow row, out error);
if (!res)
	Console.WriteLine("ERROR, unbale to read the row");
```


# Create an excel file

The code create an excel file with one sheet.

```
bool res;
ExcelError error;
ExcelProcessor proc = new ExcelProcessor();

string filename = PathFiles + "nmyExcel.xlsx";
res=proc.CreateExcelFile(filename, out ExcelFile excelFile, out error);
```


