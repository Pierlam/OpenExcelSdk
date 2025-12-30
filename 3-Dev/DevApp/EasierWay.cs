using DocumentFormat.OpenXml.Spreadsheet;
using OpenExcelSdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DevApp;

public class EasierWay
{
    public static void TestFctLight()
    {
        ExcelError error;
        bool res;
        ExcelProcessor proc = new ExcelProcessor();

        string filename = @".\Files\CellFormats.xlsx";
        ExcelFile excelFile= proc.OpenExcelFile(filename);

        ExcelSheet excelSheet= proc.GetSheetAt(excelFile, 0);

        ////--B2: int, border
        //ExcelCell excelCell = proc.GetCellAt(excelSheet, "B2");
        //var cellValueType = proc.GetCellType(excelSheet, excelCell);

        //ExcelCellValueMulti excelCellValueMulti= proc.GetCellTypeAndValue(excelSheet, excelCell);

        //excelCellValueMulti = proc.GetCellTypeAndValue(excelSheet, "B4");

        //res = proc.SetCellValue(excelSheet, "B2", new DateOnly(2025, 10, 12), "d/m/yyyy");

        proc.CloseExcelFile(excelFile);

    }
}
