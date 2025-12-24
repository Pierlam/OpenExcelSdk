# What is OpenExcelSdk ?

OpenExcelSdk is an open-source backend .NET library to use Excel (xlsx) very easily.

It's written in C#/NET8.

The only dependency is OpenXML SDK, the official Microsoft library to work with Excel files.
The last available version 3.3.0 is used.

OpenXML SDK is a big library to manage Excel, Word and Power-Point documents.
OpenExcelSdk is focus only on Excel documents.

This Microsoft library is not easy to use, so OpenExcelSdk propose a simple way to use Excel rows and cells values.
Main use cases are to get/read or set a type, define a format and set a value into a new cell or an existing one.  

OpenExcelSdk is a kind of wrapper around OpenXML SDK library.
OpenExcelSdk offers light and basic framework, main classes are : ExcelFile, ExcelSheet, ExcelRow, ExcelCell.
OpenXML SDK classes are always available in each of these classes: SpreadsheetDocument, Sheet, WorkbookPart, Sheet, Row, Cell, ...


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
string filename = @".\Files\data.xlsx";
res = proc.Open(filename, out ExcelFile excelFile, out error);

// get the first sheet of the excel file
proc.GetFirstSheet(excelFile, out ExcelSheet excelSheet, out error);

// get cell at B7 (A is 1, B is 2)
proc.GetCellAt(excelSheet, 2, 7, out ExcelCell cell, out error);

// get the type and the value of cell
proc.GetCellTypeAndValue(excelSheet, cell, out cellValueMulti, out error);

//> type of the cell can be: string, int, double, DateTime, DateOnly, TimeOnly.

// if the cell type is string, get the value and display it
if(cellValueMulti.CellType == ExcelCellType.String)
{
  Console.WriteLine("B7 type: string, value:" + proc.GetCellValueAsString(excelFile, cell));
} 
```

# Package available on Nuget

OpenExcelSdk library is packaged as a nuget ready to use:

https://www.nuget.org/packages/OpenExcelSdk

Github source repository:

https://github.com/Pierlam/OpenExcelSdk/


# Main functions

There are many available functions to get sheet, by index or by name, to read and write cell value type and format.

Take a look on ExcelProcessor main class which have many methods.

Manage basic type are: string, integer, double, DateOnly, DateTime and TimeOnly.
Set a value to a cell is possible for each of these types.

## Get a sheet

After creating of opening an excel file, the next action is to get a sheet.
There are 2 ways to get a sheet.

```
//-case1: get the first sheet of the excel file
proc.GetFirstSheet(excelFile, out ExcelSheet excelSheet, out error);

//-case 2: Get a sheet at an index, starting from 0
proc.GetSheetAt(excelFile, 0, out ExcelSheet excelSheet, out error);

//-case 3: Get a sheet by the name
proc.GetSheetByName(excelFile, "Sheet1", out ExcelSheet excelSheet, out error);
```


## Read cell value, type and format

```
// read B4 cell
proc.GetCellAt(excelSheet, 2, 4, out cell, out error);
// get the type, the format and the value
proc.GetCellTypeAndValue(excelSheet, cell, out cellValueMulti, out error);

// check: cellValueMulti.CellType, contains the type of the cell value.
// then get the value from cellValueMulti property: StringValue, IntegerValue, DoubleValue, DateOnlyValue,...
```


## Read unexisting cell 

```
// read B9 cell which not exists
proc.GetCellAt(excelSheet, 2, 9, out cell, out error);

// no error, the cell is null
if(cell==null)
{ }
```


## Set cell value

If the cell does not exists, it will be created before setting the value.

```
// set a double value into cell C10
proc.SetCellValue(excelSheet, 3, 10, 12.5, out error);

```

## Set cell format and value 

When setting a value, it's possible to define the format (display format).

```
// set a double value and format it with 2 decimals, e.g.: 12,30
proc.SetCellValue(excelSheet, 2, 9, 12.5, "0.00", out error);

// set a date with a standard format
proc.SetCellValue(excelSheet, 8, 12, new DateOnly(2025,10,12), "d/m/yyyy", out error);
```

You can use predefined format, take a look in Definitions class.
Yo can format the display of the value for number, date and currency.

For date and currency the format is mandatory.
If you set a value (string, int or double) without format in a existing cell, the defined format of the cell is used as much as possible.

```
// set a double value and format it with 2 decimals, e.g.: 12,30
proc.SetCellValue(excelSheet, 2, 5, 12.3, Definitions.NumFmtNumberTwoDec2, out error);
```

If you need another style/CellFormat, these links could you:


## Get row/last row index

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

// get row at index 0, the first one
res = proc.GetRowAt(excelSheet, 0, out ExcelRow row, out error);
if (!res)
	Console.WriteLine("ERROR, unable to read the row");
```


## Create an excel file

The code create an excel file with one sheet.

```
bool res;
ExcelError error;
ExcelProcessor proc = new ExcelProcessor();

string filename = @".\Files\data.xlsx";
res=proc.CreateExcelFile(filename, out ExcelFile excelFile, out error);
```

## Style/CellFormat/NumbergingFormat

Cell formating take an important place when read or write cell value.

In fact, manage cell value formatting for number, date and currency is always a nightmare.

The library hide this complexity to the user so you have just to set a value and a format.

Creation of a style/CellFormat is managed as better as possible.
The max number of custom style is high but limited.

Style/CellFormat are defined in the scope of a sheet, not on all the excel file.

When a new style is required, if an existing style match the new requested one, it is used in place of creating a new one.


```
// get the number of custom (user-defined) existing style/CellFormat/NumberingFormat
int count= proc.GetCustomNumberFormatsCount(excelSheet);

```

## What is not managed

The library have many functions on cell but there is several functionalities which are not managed such as: Alignment Border, Fill, Font and Protection.  


## Contact 

Have a comment to do? A question to ask? A request ?

Don't hesitate to send me an email: 

pierlam-project@outlook.com

## Resources 

### Style/CellFormat

ECMA-376, Second Edition, Part 1 - Fundamentals And Markup Language Reference section 18.8.30 page 1786:
https://www.ecma-international.org/publications-and-standards/standards/ecma-376/

Stackoverflow article on cellFormat list:
https://stackoverflow.com/questions/36670768/openxml-cell-datetype-is-null

Another Stackoverflow article on cellFormat list:
https://stackoverflow.com/questions/4655565/reading-dates-from-openxml-excel-files

### OpenXml SDK

The official web site: 
https://learn.microsoft.com/en-us/office/open-xml/open-xml-sdk

SpreadSheets, source code samples:
https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/overview