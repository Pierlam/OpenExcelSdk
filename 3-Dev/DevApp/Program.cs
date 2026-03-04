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

ExcelAllStylesExport ExportAllStyles()
{
    ExcelProcessor proc = new ExcelProcessor();

    //string filename = @"Files\CellFormat.xlsx";
    string filename = @"Files\currencies.xlsx";
    //string filename = @"Files\SetCellColorOut.xlsx";
    //string filename = @"Out\WrongSave.xlsx";


    string filenameOut = @"Out\styles.xlsx";
    //string filenameOut = @"Out\CellFormat-styles.xlsx";

    if (File.Exists(filenameOut))
        File.Delete(filenameOut);

    // export
    return proc.ExportAllStyles(filename, filenameOut);
}


void ReadCurrency()
{
    ExcelProcessor proc = new ExcelProcessor();

    string filename = @"Files\currencies.xlsx";

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

//ExcelAllStylesExport excelStyles =ExportAllStyles();

ReadCurrency();

Console.WriteLine("=> Ok, Ends.");