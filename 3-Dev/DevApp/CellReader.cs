using DocumentFormat.OpenXml.Spreadsheet;
using OpenExcelSdk;

namespace DevApp;

internal class CellReader
{
    public static void CheckFilePb()
    {
        ExcelProcessor proc = new ExcelProcessor();

        string filename = @".\Files\datLinesThenACellBlankOk.xlsx";
        ExcelFile excelFile = proc.OpenExcelFile(filename);
        ExcelSheet excelSheet= proc.GetSheetAt(excelFile, 0);

        ExcelCell cell;
        ExcelCellValue excelCellValue;
        CellFormat cellFormat;
        string dataFormat;
        StyleMgr styleMgr = new StyleMgr();

        //--A5:
        cell = proc.GetCellAt(excelSheet, "A5");
        //var cellValueType = proc.GetCellType(excelSheet, cell);
        excelCellValue= proc.GetCellValue(excelSheet, cell);

        if (excelCellValue.CellType == ExcelCellType.String)
        { }

        proc.CloseExcelFile(excelFile);
    }

    public static void ReadCellFormats()
    {
        ExcelProcessor proc = new ExcelProcessor();

        string filename = @".\Files\CellFormats.xlsx";
        ExcelFile excelFile= proc.OpenExcelFile(filename);
        ExcelSheet excelSheet = proc.GetSheetAt(excelFile, 0);

        ExcelCell cell;
        ExcelCellValue excelCellValue;
        CellFormat cellFormat;
        string dataFormat;
        StyleMgr styleMgr = new StyleMgr();

        //--B2: int, border
        cell = proc.GetCellAt(excelSheet, "B2");
        var cellValueType = proc.GetCellValue(excelSheet, cell);
        excelCellValue= proc.GetCellValue(excelSheet, cell);
        cellFormat = proc.GetCellFormat(excelSheet, cell);
        //cellFormat.BorderId

        //--B4: int, bgcolor
        // B5: red: #FF0000 // ARGB: FF + FF0000

        cell = proc.GetCellAt(excelSheet, "B5");
        excelCellValue= proc.GetCellValue(excelSheet, cell);
        cellFormat = proc.GetCellFormat(excelSheet, cell);
        if (cellFormat != null && cellFormat.BorderId != null)
        {
            uint fillId = cellFormat.FillId.Value;
            DocumentFormat.OpenXml.Spreadsheet.Fill fill = excelFile.WorkbookPart.WorkbookStylesPart.Stylesheet.Fills.ElementAt((int)fillId) as DocumentFormat.OpenXml.Spreadsheet.Fill;

            if (fill?.PatternFill?.BackgroundColor != null)
            {
                DocumentFormat.OpenXml.Spreadsheet.ForegroundColor fgColor = fill.PatternFill.ForegroundColor;

                // 2 cases: direct color or theme color

                if (fgColor.Rgb != null)
                {
                    // std yellow: "FFFFFF00"/ #FFFF00
                    Console.WriteLine($"RGB Color: {fgColor.Rgb}");
                }

                if (fgColor.Theme != null)
                {
                    Console.WriteLine("Fill color is theme-based or not set.");
                    int themeIndex = (int)fgColor.Theme.Value;
                    double tint = fgColor.Tint != null ? fgColor.Tint.Value : 0;

                    string rgb = styleMgr.GetThemeColor(excelFile.WorkbookPart, themeIndex, tint);

                    // "#70AD47"  for B4 cell
                    Console.WriteLine($"RGB Color: {rgb}");
                }
            }

            if (fill?.PatternFill?.ForegroundColor != null)
            {
                // text color
            }
        }
    }

    public static void Read()
    {
        ExcelProcessor proc = new ExcelProcessor();

        string filename = @".\Files\data.xlsx";
        ExcelFile excelFile= proc.OpenExcelFile(filename);
        ExcelSheet excelSheet = proc.GetSheetAt(excelFile, 0);

        int lastRowIdx = proc.GetLastRowIndex(excelSheet);
        Console.WriteLine("last row idx: " + lastRowIdx);

        ExcelRow row = proc.GetRowAt(excelSheet, 0);
        if (row==null)
            Console.WriteLine("ERROR, unbale to read the row");

        ExcelCell cell;

        ////--A1: values
        cell = proc.GetCellAt(excelSheet, 1, 1);
        var cellValueType = proc.GetCellValue(excelSheet, cell);
        string val = proc.GetCellValueAsString(excelSheet, cell);

        proc.CloseExcelFile(excelFile);
    }
}