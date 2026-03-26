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


void DataTableHasEmptyRow()
{

    ExcelProcessor proc = new ExcelProcessor();

    // open an excel file
    string filename = @"Files\datatableHasEmptyRow.xlsx";
    ExcelFile excelFile = proc.OpenExcelFile(filename);

    // get the first sheet of the excel file
    ExcelSheet excelSheet = proc.GetFirstSheet(excelFile);

    int lastRowIdx = proc.GetLastRowIndex(excelSheet);
    Console.WriteLine($"LastRowIndex: {lastRowIdx}");

    // Row #1 to Row #6: row #3 and row #5 are empty

    int rowIndex = -1;

    List <ExcelCell> listCells = proc.GetRowCellsAtAddress(excelSheet, lastRowIdx);
    if (listCells.Count > 0)
    {
        ExcelCellAddressUtils.GetColumnAndRowAddress(listCells[0].Cell.CellReference.Value, out int colIndex, out rowIndex);
    }


    //--scan each row
    for (int i = 1; i <= lastRowIdx; i++)
    {
        Console.WriteLine($"---");

        // get the row by index, if the row doesn't exists, row is null, it's not an error
        ExcelRow excelRow = proc.GetRowAtIndex(excelSheet, i);

        // NO! dont't  scan cells like this wth GetRowAt
        // get first cell of the row
        ExcelCell excelCell = proc.GetCellAt(excelSheet,1,i); 
        if (excelCell == null) 
        {
            Console.WriteLine($"Row #{i}, Num:{rowIndex}, Cell is null.");
            continue;
        }
        Console.WriteLine($"Row #{i}, Num:{rowIndex}, Cell {excelCell.Cell.CellReference}, has a value");

    }
}

// 4 Rows defined. lastRowAddress=6, RowAddr=3 and 5 are empty.
void ScanDataTableWayOne()
{

    ExcelProcessor proc = new ExcelProcessor();

    Console.WriteLine("ScanDataTableWayOne: Does NOT display empty rows and cells!");

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
        // get the row by index, if the row doesn't exists, row is null, it's not an error
        ExcelRow excelRow = proc.GetRowAtIndex(excelSheet, r);

        // get cells of the row
        List<ExcelCell> listCells = proc.GetRowCells(excelSheet, excelRow);

        // scan each cell of the row
        foreach (ExcelCell cell in listCells) 
        {
            Console.WriteLine($"Row #{r}, Cell addr: {cell.Cell.CellReference}");
        }
    }

    /*
=> OpenExcelSdk DevApp:
LastRowIndex: 4
Row #1, Cell addr: A1
Row #2, Cell addr: A2
Row #3, Cell addr: A4
Row #4, Cell addr: A6
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

    /*
     */
}

// 4 Rows defined. lastRowAddress=6, RowAddr=3 and 5 are empty.
void ScanDataTableWayThree()
{

    ExcelProcessor proc = new ExcelProcessor();

    Console.WriteLine("ScanDataTableWayTwo: Display empty rows and cells");

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
                Console.WriteLine($"Cell addr: Col:{c}, Row{r} is empty");
            }
            else
            {
                Console.WriteLine($"Cell addr: {cell.Cell.CellReference}, has a value");
            }
        }
    }

    /*
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

//ScanDataTableWayOne();
//ScanDataTableWayOneOne();
ScanDataTableWayTwo();

//ExcelAllStylesExport excelStyles =ExportAllStyles();

//ReadCurrency();

Console.WriteLine("=> Ok, Ends.");