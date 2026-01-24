using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenExcelSdk.System;
using OpenExcelSdk.System.Export;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcelSdk.Export;


public class ExcelStylesExporter
{

    public static void Export(ExcelProcessor excelProcessor, ExcelAllStylesExport excelStyles, ExcelFile excelFileOut)
    {
        ExcelSheet excelSheetOut = excelProcessor.CreateSheet(excelFileOut, "Styles - CellFormat");

        // create the out header
        CreateOutHeader(excelProcessor, excelSheetOut);

        int i = 0;
        foreach(ExcelStyleExport styleExport in excelStyles.ListStyles)
        {
            string rowIdx = (i + 2).ToString();

            excelProcessor.SetCellValue(excelSheetOut, "A" + rowIdx, styleExport.SheetIndex);
            excelProcessor.SetCellValue(excelSheetOut, "B" + rowIdx, excelStyles.ListSheets.FirstOrDefault(s => s.Index== styleExport.SheetIndex).Name);
            excelProcessor.SetCellValue(excelSheetOut, "C" + rowIdx, i);
            excelProcessor.SetCellValue(excelSheetOut, "D" + rowIdx, styleExport.NumberFormatId);
            excelProcessor.SetCellValue(excelSheetOut, "E" + rowIdx, styleExport.NumberFormat);

            //--K: currency
            Currency? currency = CurrencyMgr.CreateCurrency(styleExport.NumberFormat);
            if (currency != null)
                excelProcessor.SetCellValue(excelSheetOut, "F" + rowIdx, currency.Code + "-" + currency.Name);

            excelProcessor.SetCellValue(excelSheetOut, "G" + rowIdx, styleExport.FillId);

            ExcelFillExport fillExport = excelStyles.ListFills.FirstOrDefault(f => f.FillId == styleExport.FillId);

            if (fillExport == null)
            {
                excelProcessor.SetCellValue(excelSheetOut, "H" + rowIdx, "ERROR");
            }
            else
            {

                if (fillExport.FgColor != null)
                {
                    excelProcessor.SetCellValue(excelSheetOut, "H" + rowIdx, fillExport.FgColor.Rgb);
                }
            }

            excelProcessor.SetCellValue(excelSheetOut, "I" + rowIdx, styleExport.BorderId);
            excelProcessor.SetCellValue(excelSheetOut, "J" + rowIdx, styleExport.FontId);

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
        proc.SetCellValue(excelSheet, "C1", "StyleIndex");
        proc.SetCellValue(excelSheet, "D1", "NumberFormatId");
        proc.SetCellValue(excelSheet, "E1", "NumberFormat");
        proc.SetCellValue(excelSheet, "F1", "Currency");
        proc.SetCellValue(excelSheet, "G1", "FillId");
        proc.SetCellValue(excelSheet, "H1", "Fill.FgColor");
        proc.SetCellValue(excelSheet, "I1", "BorderId");
        proc.SetCellValue(excelSheet, "J1", "FontId");
    }
}
