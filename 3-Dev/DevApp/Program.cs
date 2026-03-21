using DevApp;
using DocumentFormat.OpenXml;
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

    // Row #1 to Row #5: row #3 is empty

    //--scan each row
    for (int i = 1; i <= lastRowIdx; i++)
    {
        Console.WriteLine($"---");

        // get the row by index, if the row doesn't exists, row is null, it's not an error
        ExcelRow excelRow = proc.GetRowAt(excelSheet, i);

        int rowIndex = -1;

        List<ExcelCell> listCells= proc.GetRowCells(excelSheet, i);
        if (listCells.Count > 0)
        {
            ExcelCellAddressUtils.GetColumnAndRowIndex(listCells[0].Cell.CellReference.Value, out int colIndex, out rowIndex);
        }

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

DataTableHasEmptyRow();

//ExcelAllStylesExport excelStyles =ExportAllStyles();

//ReadCurrency();

Console.WriteLine("=> Ok, Ends.");