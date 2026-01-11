using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenExcelSdk.System.Export;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcelSdk;


public class StylesExporter
{
    public static void Export(ExcelProcessor excelProcessor, ExcelStyles excelStyles, string filenameOut)
    {
        ExcelFile excelFileOut = excelProcessor.CreateExcelFile(filenameOut);
        ExcelSheet excelSheetOut = excelProcessor.GetSheetAt(excelFileOut, 0);

        // create the out header
        CreateOutHeader(excelProcessor, excelSheetOut);

        int i = 0;
        foreach(StyleTable styleTable in excelStyles.ListStyleTable)
        {
            string rowIdx = (i + 2).ToString();

            excelProcessor.SetCellValue(excelSheetOut, "A" + rowIdx, styleTable.SheetIndex);
            excelProcessor.SetCellValue(excelSheetOut, "B" + rowIdx, excelStyles.ListSheetTable.FirstOrDefault(s => s.Index== styleTable.SheetIndex).Name);
            excelProcessor.SetCellValue(excelSheetOut, "C" + rowIdx, i);
            excelProcessor.SetCellValue(excelSheetOut, "D" + rowIdx, styleTable.NumberFormatId);
            excelProcessor.SetCellValue(excelSheetOut, "E" + rowIdx, styleTable.NumberFormat);
            excelProcessor.SetCellValue(excelSheetOut, "F" + rowIdx, styleTable.FillId);
            excelProcessor.SetCellValue(excelSheetOut, "G" + rowIdx, styleTable.FillPattern);

            if (styleTable.BgColor != null)
            {
                if (styleTable.BgColor.ThemeIndex > 0)
                {
                    excelProcessor.SetCellValue(excelSheetOut, "H" + rowIdx, styleTable.BgColor.ThemeIndex);
                    excelProcessor.SetCellValue(excelSheetOut, "I" + rowIdx, styleTable.BgColor.Rgb);
                }else
                    excelProcessor.SetCellValue(excelSheetOut, "I" + rowIdx, styleTable.BgColor.ARgb);
            }

            if (styleTable.FgColor != null)
            {
                if (styleTable.FgColor.ThemeIndex > 0)
                {
                    excelProcessor.SetCellValue(excelSheetOut, "J" + rowIdx, styleTable.FgColor.ThemeIndex);
                    excelProcessor.SetCellValue(excelSheetOut, "K" + rowIdx, styleTable.FgColor.Rgb);
                }
                else
                    excelProcessor.SetCellValue(excelSheetOut, "K" + rowIdx, styleTable.FgColor.ARgb);
            }

            i++;
        }

        excelProcessor.CloseExcelFile(excelFileOut);

    }

    // create the out header
    static void CreateOutHeader(ExcelProcessor proc, ExcelSheet excelSheet)
    {
        proc.SetCellValue(excelSheet, "A1", "SheetIdx");
        proc.SetCellValue(excelSheet, "B1", "SheetName");
        proc.SetCellValue(excelSheet, "C1", "StyleIndex");
        proc.SetCellValue(excelSheet, "D1", "NumberFormatId");
        proc.SetCellValue(excelSheet, "E1", "NumberFormat");
        proc.SetCellValue(excelSheet, "F1", "FillId");
        proc.SetCellValue(excelSheet, "G1", "Fill.PatternType");
        proc.SetCellValue(excelSheet, "H1", "Fill.BgColor.ThemeIdx");
        proc.SetCellValue(excelSheet, "I1", "Fill.BgColor");
        proc.SetCellValue(excelSheet, "J1", "Fill.FgColor.ThemeIdx");
        proc.SetCellValue(excelSheet, "K1", "Fill.FgColor");
    }


}
