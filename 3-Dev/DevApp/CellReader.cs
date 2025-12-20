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
    public static void Read()
    {
        ExcelProcessor proc = new ExcelProcessor();

        //string filename = @".\Files\myexcel.xlsx";
        //proc.CreateSpreadsheetWorkbook(filename);

        ExcelError error;
        bool res;

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

        ////--A2: 12
        //res = proc.GetCellAt(excelSheet, 1, 2, out cell, out error);
        //cellValueType = proc.GetCellType(excelSheet, cell);
        //proc.GetCellValue(excelSheet, cell, out int intValue, out error);

        ////--A3: 56,67
        //res = proc.GetCellAt(excelSheet, 1, 3, out cell, out error);
        //cellValueType = proc.GetCellType(excelSheet, cell);
        //proc.GetCellValue(excelSheet, cell, out double doubleValue, out error);

        ////--A4: is null
        //res = proc.GetCellAt(excelSheet, 1, 4, out cell, out error);
        //proc.CreateCell(excelSheet, 1,4, out cell, out error);
        //proc.SetCellValue(excelSheet, cell, "hello", out error);
        //proc.GetCellValue(excelSheet, cell, out stringValue, out error);

        ////--A5: is null -> 17
        //proc.CreateCell(excelSheet, 1, 5, out cell, out error);
        //proc.SetCellValue(excelSheet, cell, 17, out error);
        //proc.GetCellValue(excelSheet, cell, out intValue, out error);

        ////--A6: is null -> 23.45
        //proc.CreateCell(excelSheet, 1, 6, out cell, out error);
        //proc.SetCellValue(excelSheet, cell, 23.45, out error);
        //proc.GetCellValue(excelSheet, cell, out doubleValue, out error);

        //// set cell value
        //// date,...

        //// update the value of an existing cell


        if (!proc.Close(excelFile, out error))
            Console.WriteLine("ERROR, Unable to close the Excel file.");

    }
}
