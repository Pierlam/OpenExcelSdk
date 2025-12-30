using DevApp;
using OpenExcelSdk;

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

Console.WriteLine("=> OpenExcelSdk DevApp:");

//CellReader.Read();

//ConvertDouble();

//DevCloneStyle();

//CellReader.ReadCellFormats();

//CellReader.CheckFilePb();

EasierWay.TestFctLight();

Console.WriteLine("=> Ok, Ends.");