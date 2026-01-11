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

ExcelStyles ExportStyles()
{
    ExcelProcessor proc = new ExcelProcessor();
    //string filename = @"Files\DevCloneStyle.xlsx";
    string filename = @"Files\CellFormats.xlsx";
    string filenameOut = @"Out\ListStyles.xlsx";

    if (File.Exists(filenameOut))
        File.Delete(filenameOut);

    // export
    return proc.ExportStyles(filename, filenameOut);
}

Console.WriteLine("=> OpenExcelSdk DevApp:");

//CellReader.Read();

//ConvertDouble();

//DevCloneStyle();


//CellReader.CheckFilePb();

//EasierWay.TestFctLight();

//CellReader.ReadCellFormats();


//var Rgb = HexBinaryValue.FromString("#00FF00");
ExcelStyles excelStyles =ExportStyles();

Console.WriteLine("=> Ok, Ends.");