using DevApp;
using DocumentFormat.OpenXml.Math;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenExcelSdk;
using OpenExcelSdk.System;

void DevCloneStyle()
{
    ExcelProcessor proc = new ExcelProcessor();
    bool res;
    ExcelError error;
    string filename = @"Files\DevCloneStyle.xlsx";

    res = proc.Open(filename, out ExcelFile excelFile, out error);
    res = proc.GetSheetAt(excelFile, 0, out ExcelSheet excelSheet, out error);

    //--B2: dateTime, custom, BgColor, FgCOlor, Border: 09/12/2021 12:30:45
    res = proc.GetCellAt(excelSheet, 2, 2, out ExcelCell excelCell, out error);
    proc.SetCellValue(excelSheet, excelCell, "Bonjour", out error);

    //StyleMgr styleMgr = new StyleMgr();
    //res = styleMgr.CloneStyle(excelSheet, excelCell);
    //if (!res) return;
    //proc.Close(excelFile, out error);


    // save the changes
    res = proc.Close(excelFile, out error);
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


CellReader.ReadCellFormats();

Console.WriteLine("=> Ok, Ends." );
