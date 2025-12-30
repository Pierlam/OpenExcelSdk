# What is OpenExcelSdk ?

OpenExcelSdk is an open-source backend .NET library to use Excel (xlsx) very easily.

It's written in C#/NET8 with VS2026. The code is covered by 18 unit tests to ensure the non-regression of incoming evolutions.

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
excelCellValue excelCellValue;

// create the excel processor to read/create/update cell
ExcelProcessor proc = new ExcelProcessor();

// open an excel file
string filename = @".\Files\data.xlsx";
ExcelFile excelFile = proc.OpenExcelFile(filename);

// get the first sheet of the excel file
ExcelSheet excelSheet = proc.GetFirstSheet(excelFile);

// get cell at B7, if the cell doesn't exists, cell is null, it's not an error 
ExcelCell cell= proc.GetCellAt(excelSheet, "B7");

// get the type and the value of cell
excelCellValue= proc.GetCellValue(excelSheet, cell);

//> type of the cell can be: string, int, double, DateTime, DateOnly, TimeOnly.

// if the cell type is string, get the value and display it
if(excelCellValue.CellType == ExcelCellType.String)
{
  Console.WriteLine("B7 type: string, value:" + proc.GetCellValueAsString(excelFile, cell));
} 
```

# Package available on Nuget

OpenExcelSdk library is packaged as a nuget ready to use:

https://www.nuget.org/packages/OpenExcelSdk

Github source repository:

https://github.com/Pierlam/OpenExcelSdk/

Github wiki:

https://github.com/Pierlam/OpenExcelSdk/wiki


# Main functions

There are many available functions to get sheet, by index or by name, to read and write cell value type and format.

Take a look on ExcelProcessor main class which have many methods.

Manage basic type are: string, integer, double, DateOnly, DateTime and TimeOnly.
Set a value to a cell is possible for each of these types.

## Get a sheet

After creating or opening an excel file, the next action is to get a sheet.
There are 3 ways to get a sheet.

```
//-case1: get the first sheet of the excel file
ExcelSheet excelSheet= proc.GetFirstSheet(excelFile);

//-case 2: Get a sheet at an index, starting from 0
ExcelSheet excelSheet= proc.GetSheetAt(excelFile, 0);

//-case 3: Get a sheet by the name
ExcelSheet excelSheet= proc.GetSheetByName(excelFile, "Sheet1");
```


## Get cell value, type and format

```
// read B4 cell
ExcelCell excelCell= proc.GetCellAt(excelSheet, "B4");

// or like this:  ExcelCell excelCell= proc.GetCellAt(excelSheet, 2,4);

// get the type, the format and the value
excelCellValue excelCellValue= proc.GetCellValue(excelSheet, excelCell);

// check: excelCellValue.CellType, contains the type of the cell value.
// then get the value from excelCellValue property: StringValue, IntegerValue, DoubleValue, DateOnlyValue and TimeOnly

// the cell value can be empty/Blank, in some cases the type will be undefined
if(excelCellValue.IsEmpty) ...

```


## Read unexisting cell 

```
// read B9 cell which not exists
ExcelCell excelCell= proc.GetCellAt(excelSheet, "B9"2, 9");

// the cell is null, it's not an error
if(excelCell==null)
{ // do something }
```


## Set cell value

If the cell does not exists, it will be created before setting the value. If the cell contains a formula, it will be revoved.

```
// set a double value into cell C10
proc.SetCellValue(excelSheet, "C10", 12.5);

// or 
proc.SetCellValue(excelSheet, 3, 10, 12.5);

```

## Set cell format and value 

When setting a value, it's possible to define the format (display format).
You can format the display of the value for number, date and currency.

For date and currency the format is mandatory.

Example, format the display of a number:

```
// set a double value and format it with 2 decimals, e.g.: 12,30
proc.SetCellValue(excelSheet, "B9", 12.5, "0.00");

// set a date with a standard format
proc.SetCellValue(excelSheet, "D12", new DateOnly(2025,10,12), "d/m/yyyy");
```

You can use one of some predefined format declared in the class Definitions.cs.


```
// set a double value and format it with 2 decimals, e.g.: 12,30
// Definitions.NumFmtNumberTwoDec2= "0.00"
proc.SetCellValue(excelSheet, "B5", 12.3, Definitions.NumFmtNumberTwoDec2);

// set a formated date
// Definitions.NumFmtDayMonthYear14= "d/m/yyyy"
proc.SetCellValue(excelSheet, "C4", new DateOnly(2025,12,28), Definitions.NumFmtDayMonthYear14);
```

If you set a value (string, int or double) without format in a existing cell, the defined format of the cell is used as much as possible.


## Get row/last row index

The code below get a the last row index, and also get a row at an index.

```
bool res;
ExcelProcessor proc = new ExcelProcessor();

// open an existing excel file
string filename = @".\Files\data.xlsx";
ExcelFile excelFile= proc.OpenExcelFile(filename);

// get the first sheet
ExcelSheet excelSheet= proc.GetSheetAt(excelFile, 0);

// get the index of the last row containing cells
int lastRowIdx = proc.GetLastRowIndex(excelSheet);
Console.WriteLine("last row idx: " + lastRowIdx);

// get the row at index 0, the first one
ExcelRow row = proc.GetRowAt(excelSheet, 0);
if (row==null)
	Console.WriteLine("ERROR, unable to read the row");
```


## Create an excel file

The code below will create an excel file with one sheet. 

```
bool res;
ExcelProcessor proc = new ExcelProcessor();

// create an excel with one sheet, the name will: Sheet1
string filename = @".\Files\data.xlsx";
ExcelFile excelFile= proc.CreateExcelFile(filename);

// or set the name of the first sheet
ExcelFile excelFile= proc.CreateExcelFile(filename, "MySheet");

```

## Style/CellFormat/NumbergingFormat

Cell formating take an important place when read or write cell value.

Technically, Excel has two kind of format, built-in and custom.
built-in are only identified by dedicated id, from 0 to 163.

Custom format are defined by a string which represents the format.

Manage cell value formatting for number, date and currency is a nightmare.

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

The library offers many functions on cell but there are several functionalities which are not managed such as: Alignment Border, Fill, Font and Protection.  

But the background and foreground color (Fill) get/set will be the next feature implemented. 


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