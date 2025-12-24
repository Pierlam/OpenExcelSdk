using DocumentFormat.OpenXml.Spreadsheet;
using OpenExcelSdk;
using OpenExcelSdk.System;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DevApp;
internal class CellReader
{
    public static void ReadCellFormats()
    {
        ExcelError error;
        bool res;
        ExcelProcessor proc = new ExcelProcessor();

        string filename = @".\Files\CellFormats.xlsx";
        proc.Open(filename, out ExcelFile excelFile, out error);
        proc.GetSheetAt(excelFile, 0, out ExcelSheet excelSheet, out error);

        ExcelCell cell;
        ExcelCellValueMulti cellValueMulti;
        CellFormat cellFormat;
        string dataFormat;
        StyleMgr styleMgr = new StyleMgr();

        //--A1: 
        res = proc.GetCellAt(excelSheet, 1, 1, out cell, out error);
        var cellValueType = proc.GetCellType(excelSheet, cell);
        proc.GetCellTypeAndValue(excelSheet, cell, out cellValueMulti, out error);

        cellFormat = proc.GetCellFormat(excelSheet, cell);
        // 21: built-in format: HH:mm:ss
        int fmtId= (int)cellFormat.NumberFormatId.Value;


    }

    public static void Read()
    {
        ExcelError error;
        bool res;
        ExcelProcessor proc = new ExcelProcessor();

        string filename = @".\Files\data.xlsx";
        proc.Open(filename, out ExcelFile excelFile, out error);
        proc.GetSheetAt(excelFile, 0, out ExcelSheet excelSheet, out error);

        int lastRowIdx = proc.GetLastRowIndex(excelSheet);
        Console.WriteLine("last row idx: " + lastRowIdx);

        res = proc.GetRowAt(excelSheet, 0, out ExcelRow row, out error);
        if (!res)
            Console.WriteLine("ERROR, unbale to read the row");

        ExcelCell cell;

        ////--A1: values
        res = proc.GetCellAt(excelSheet, 1, 1, out cell, out error);
        var cellValueType = proc.GetCellType(excelSheet, cell);
        string val = proc.GetCellValueAsString(excelSheet, cell);



        if (!proc.Close(excelFile, out error))
            Console.WriteLine("ERROR, Unable to close the Excel file.");

    }
}
