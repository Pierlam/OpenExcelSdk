using OpenExcelSdk.System.Export;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcelSdk.Export;

public class ExcelFillsExporter
{
    // tabpage two: Fills
    public static void ExportFills(ExcelProcessor excelProcessor, ExcelAllStylesExport excelStyles, ExcelFile excelFileOut)
    {
        ExcelSheet excelSheetOut = excelProcessor.CreateSheet(excelFileOut, "Fills");

        // create the out header
        CreateOutHeader(excelProcessor, excelSheetOut);

        int i = 0;
        foreach (ExcelFillExport fillExport in excelStyles.ListFills)
        {
            string rowIdx = (i + 2).ToString();

            excelProcessor.SetCellValue(excelSheetOut, "A" + rowIdx, fillExport.SheetIndex);
            excelProcessor.SetCellValue(excelSheetOut, "B" + rowIdx, excelStyles.ListSheet.FirstOrDefault(s => s.Index == fillExport.SheetIndex).Name);
            excelProcessor.SetCellValue(excelSheetOut, "C" + rowIdx, fillExport.FillId);
            excelProcessor.SetCellValue(excelSheetOut, "D" + rowIdx, fillExport.PatternType);

            if (fillExport.BgColor != null)
            {
                if (fillExport.BgColor.ThemeIndex > 0)
                    excelProcessor.SetCellValue(excelSheetOut, "E" + rowIdx, fillExport.BgColor.ThemeIndex);

                excelProcessor.SetCellValue(excelSheetOut, "F" + rowIdx, fillExport.BgColor.Rgb);
            }

            if (fillExport.FgColor != null)
            {
                if (fillExport.FgColor.ThemeIndex > 0)
                    excelProcessor.SetCellValue(excelSheetOut, "G" + rowIdx, fillExport.FgColor.ThemeIndex);

                excelProcessor.SetCellValue(excelSheetOut, "H" + rowIdx, fillExport.FgColor.Rgb);
            }

            excelProcessor.SetCellValue(excelSheetOut, "I" + rowIdx, fillExport.ListGradient.Count);

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
        proc.SetCellValue(excelSheet, "C1", "FillId");
        proc.SetCellValue(excelSheet, "D1", "PatternType");
        proc.SetCellValue(excelSheet, "E1", "BgColor.ThemeIdx");
        proc.SetCellValue(excelSheet, "F1", "BgColor");
        proc.SetCellValue(excelSheet, "G1", "FgColor.ThemeIdx");
        proc.SetCellValue(excelSheet, "H1", "FgColor");
        proc.SetCellValue(excelSheet, "I1", "NbGradient");
    }

}

