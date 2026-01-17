using OpenExcelSdk.System.Export;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcelSdk.Export;

public class ExcelFontsExporter
{
    public static void ExportFonts(ExcelProcessor excelProcessor, ExcelAllStylesExport excelStyles, ExcelFile excelFileOut)
    {
        ExcelSheet excelSheetOut = excelProcessor.CreateSheet(excelFileOut, "Fonts");

        // create the out header
        CreateOutHeader(excelProcessor, excelSheetOut);

        int i = 0;
        foreach (ExcelFontExport fontExport in excelStyles.ListFonts)
        {
            string rowIdx = (i + 2).ToString();

            excelProcessor.SetCellValue(excelSheetOut, "A" + rowIdx, fontExport.SheetIndex);
            excelProcessor.SetCellValue(excelSheetOut, "B" + rowIdx, excelStyles.ListSheets.FirstOrDefault(s => s.Index == fontExport.SheetIndex).Name);
            excelProcessor.SetCellValue(excelSheetOut, "C" + rowIdx, fontExport.FontId);

            i++;
        }
    }

    /// <summary>
    /// Create the out header
    /// </summary>
    /// <param name="proc"></param>
    /// <param name="excelSheet"></param>
    static void CreateOutHeader(ExcelProcessor proc, ExcelSheet excelSheet)
    {
        proc.SetCellValue(excelSheet, "A1", "SheetIdx");
        proc.SetCellValue(excelSheet, "B1", "SheetName");
        proc.SetCellValue(excelSheet, "C1", "FontId");
    }

}