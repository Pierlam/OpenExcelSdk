using DevApp;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenExcelSdk;
using OpenExcelSdk.System.Export;

void DevCloneStyle()
{
    ExcelProcessor proc = new ExcelProcessor();
    string filename = @"Files\DevCloneStyle.xlsx";

    ExcelFile excelFile = proc.OpenExcelFile(filename);
    ExcelSheet excelSheet = proc.GetSheetAt(excelFile, 0);

    //--B2: dateTime, custom, BgColor, FgCOlor, Border: 09/12/2021 12:30:45
    ExcelCell excelCell = proc.GetCellAt(excelSheet, 2, 2);
    proc.SetCellValue(excelSheet, excelCell, "Bonjour");

    //StyleMgr styleMgr = new StyleMgr();
    //res = styleMgr.CloneStyle(excelSheet, excelCell);
    //if (!res) return;
    //proc.Close(excelFile, out error);

    // save the changes
    proc.CloseExcelFile(excelFile);
}

void ConvertDouble()
{
    string value = "45927.524259259262";
    value = value.Replace('.', ',');

    //string value = "45927,524";
    double valDouble = double.Parse(value);
}


// 4 Rows defined. lastRowAddress=6, RowAddr=3 and 5 are empty.
void ScanDataTableWayOne()
{

    ExcelProcessor proc = new ExcelProcessor();

    Console.WriteLine("ScanDataTableWayOne: Scan only existing rows and cells");

    // open an excel file
    string filename = @"Files\scanDatatable.xlsx";
    ExcelFile excelFile = proc.OpenExcelFile(filename);

    // get the first sheet of the excel file
    ExcelSheet excelSheet = proc.GetFirstSheet(excelFile);

    int lastRowIdx = proc.GetLastRowIndex(excelSheet);
    Console.WriteLine($"LastRowIndex: {lastRowIdx}");


    //--scan each existing row
    for (int r = 1; r <= lastRowIdx; r++)
    {
        Console.WriteLine("---");
        Console.WriteLine($"Row idx:{r}");

        // get the row by index, if the row doesn't exists, row is null, it's not an error
        ExcelRow excelRow = proc.GetRowAtIndex(excelSheet, r);

        // get cells of the row
        List<ExcelCell> listCells = proc.GetRowCells(excelSheet, excelRow);

        // scan each cell of the row
        foreach (ExcelCell cell in listCells) 
        {
            Console.WriteLine($"Cell addr: {cell.Cell.CellReference} has a value");
        }
    }

    proc.CloseExcelFile(excelFile);

    /*
=> OpenExcelSdk DevApp:
ScanDataTableWayOne: Does NOT display empty rows and cells!
LastRowIndex: 4
---
Row idx:1
Cell addr: A1 has a value
Cell addr: B1 has a value
Cell addr: C1 has a value
---
Row idx:2
Cell addr: A2 has a value
Cell addr: B2 has a value
Cell addr: C2 has a value
---
Row idx:3
Cell addr: A4 has a value
Cell addr: C4 has a value
---
Row idx:4
Cell addr: A6 has a value
Cell addr: B6 has a value
=> Ok, Ends.
    */
}


// 4 Rows defined. lastRowAddress=6, RowAddr=3 and 5 are empty.
void ScanDataTableWayTwo()
{

    ExcelProcessor proc = new ExcelProcessor();

    Console.WriteLine("ScanDataTableWayTwo: Display empty rows but not empty cells");

    // open an excel file
    string filename = @"Files\scanDatatable.xlsx";
    ExcelFile excelFile = proc.OpenExcelFile(filename);

    // get the first sheet of the excel file
    ExcelSheet excelSheet = proc.GetFirstSheet(excelFile);

    int lastRowAddr = proc.GetLastRowAddress(excelSheet);
    Console.WriteLine($"LastRowAddress: {lastRowAddr}");


    //--scan each existing row
    for (int r = 1; r <= lastRowAddr; r++)
    {
        Console.WriteLine("---");
        Console.WriteLine($"Row addr:{r}");

        // get cells of the row
        List<ExcelCell> listCells = proc.GetRowCellsAtAddress(excelSheet, r);

        // scan each cell of the row
        foreach (ExcelCell cell in listCells)
        {
            Console.WriteLine($"Cell addr: {cell.Cell.CellReference}");
        }
    }

    proc.CloseExcelFile(excelFile);

    /*
     * ScanDataTableWayTwo: Display empty rows but not empty cells
LastRowAddress: 6
---
Row addr:1
Cell addr: A1
Cell addr: B1
Cell addr: C1
---
Row addr:2
Cell addr: A2
Cell addr: B2
Cell addr: C2
---
Row addr:3
---
Row addr:4
Cell addr: A4
Cell addr: C4
---
Row addr:5
---
Row addr:6
Cell addr: A6
Cell addr: B6
 */
}

// 4 Rows defined. lastRowAddress=6, RowAddr=3 and 5 are empty.
void ScanDataTableWayThree()
{

    ExcelProcessor proc = new ExcelProcessor();

    Console.WriteLine("ScanDataTableWayThree: Display empty rows and cells");

    // open an excel file
    string filename = @"Files\scanDatatable.xlsx";
    ExcelFile excelFile = proc.OpenExcelFile(filename);

    // get the first sheet of the excel file
    ExcelSheet excelSheet = proc.GetFirstSheet(excelFile);

    int lastRowAddr = proc.GetLastRowAddress(excelSheet);
    Console.WriteLine($"LastRowAddress: {lastRowAddr}");


    //--scan each existing row
    for (int r = 1; r <= lastRowAddr; r++)
    {
        Console.WriteLine("---");
        Console.WriteLine($"Row addr:{r}");

        int lastColAddr = proc.GetLastColAddress(excelSheet, r);

        for(int c = 1; c <= lastColAddr; c++)
        {
            ExcelCell cell = proc.GetCellAt(excelSheet, c, r);
            if (cell == null)
            {
                Console.WriteLine($"Cell addr: Col:{c}, Row{r}: cell is empty");
            }
            else
            {
                Console.WriteLine($"Cell addr: {cell.Cell.CellReference}: Cell has a value");
            }
        }
    }

    proc.CloseExcelFile(excelFile);

    /*
ScanDataTableWayThree: Display empty rows and cells
LastRowAddress: 6
---
Row addr:1
Cell addr: A1: Cell has a value
Cell addr: B1: Cell has a value
Cell addr: C1: Cell has a value
---
Row addr:2
Cell addr: A2: Cell has a value
Cell addr: B2: Cell has a value
Cell addr: C2: Cell has a value
---
Row addr:3
---
Row addr:4
Cell addr: A4: Cell has a value
Cell addr: Col:2, Row4: cell is empty
Cell addr: C4: Cell has a value
---
Row addr:5
---
Row addr:6
Cell addr: A6: Cell has a value
Cell addr: B6: Cell has a value
=> Ok, Ends.      
     */
}



ExcelAllStylesExport ExportAllStyles()
{
    ExcelProcessor proc = new ExcelProcessor();

    //string filename = @"Files\CellFormat.xlsx";
    string filename = @"Files\currencies.xlsx";


    //string filename = @"Files\currencyAccounting.xlsx";

    //string filename = @"Files\SetCellValueCurrency_pb_accountingUS.xlsx";

    //SetCellValueCurrency_Empty
    //string filename = @"Files\SetCellValueCurrency_Empty.xlsx";


    //string filename = @"Files\SetCellColorOut.xlsx";
    //string filename = @"Out\WrongSave.xlsx";


    string filenameOut = @"Out\styles.xlsx";
    //string filenameOut = @"Out\CellFormat-styles.xlsx";


    Console.WriteLine("=> ExportAllStyles, file: "  +filename);

    if (File.Exists(filenameOut))
        File.Delete(filenameOut);

    // export
    return proc.ExportAllStyles(filename, filenameOut);
}


void ReadCurrency()
{
    ExcelProcessor proc = new ExcelProcessor();

    string filename = @"Files\currencyAccounting.xlsx";

    ExcelFile excelFile = proc.OpenExcelFile(filename);
    ExcelSheet excelSheet = proc.GetFirstSheet(excelFile);
    ExcelCellValue cellValue = proc.GetCellValue(excelSheet, "B2");

    proc.CloseExcelFile(excelFile);
}

void CreateWrongExcel()
{
    ExcelProcessor proc = new ExcelProcessor();

    string filename = @"Out\Wrong.xlsx";

    if (File.Exists(filename))
        File.Delete(filename);

    ExcelFile excelFile = proc.CreateExcelFile(filename);

    ExcelSheet excelSheet = proc.GetFirstSheet(excelFile);
    proc.SetCellValue(excelSheet, "B3", 34);

    proc.SetCellValue(excelSheet, "A2+", 34);


    proc.CloseExcelFile(excelFile);
}


Console.WriteLine("=> OpenExcelSdk DevApp:");

//CellReader.Read();

//ConvertDouble();

//DevCloneStyle();


//CellReader.CheckFilePb();

//EasierWay.TestFctLight();

//CellReader.ReadCellFormats();


//var Rgb = HexBinaryValue.FromString("#00FF00");


//CreateWrongExcel();

//DataTableHasEmptyRow();

ScanDataTableWayOne();
//ScanDataTableWayTwo();
//ScanDataTableWayThree();

//ExcelAllStylesExport excelStyles =ExportAllStyles();

//ReadCurrency();

Console.WriteLine("=> Ok, Ends.");