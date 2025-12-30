using OpenExcelSdk;

namespace DevApp;

public class EasierWay
{
    public static void TestFctLight()
    {
        ExcelProcessor proc = new ExcelProcessor();

        string filename = @".\Files\CellFormats.xlsx";
        ExcelFile excelFile = proc.OpenExcelFile(filename);

        ExcelSheet excelSheet = proc.GetSheetAt(excelFile, 0);

        ////--B2: int, border
        ExcelCell excelCell = proc.GetCellAt(excelSheet, "B2");
        //var cellValueType = proc.GetCellType(excelSheet, excelCell);

        //excelCellValue excelCellValue= proc.GetCellTypeAndValue(excelSheet, excelCell);

        //excelCellValue = proc.GetCellTypeAndValue(excelSheet, "B4");

        //res = proc.SetCellValue(excelSheet, "B2", new DateOnly(2025, 10, 12), "d/m/yyyy");

        proc.CloseExcelFile(excelFile);
    }
}