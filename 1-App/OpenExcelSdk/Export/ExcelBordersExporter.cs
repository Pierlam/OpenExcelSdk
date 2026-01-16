using OpenExcelSdk.System.Export;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcelSdk.Export;

public class ExcelBordersExporter
{
    public static void ExportBorders(ExcelProcessor excelProcessor, ExcelAllStylesExport excelStyles, ExcelFile excelFileOut)
    {
        ExcelSheet excelSheetOut = excelProcessor.CreateSheet(excelFileOut, "Borders");

        // create the out header
        CreateOutHeader(excelProcessor, excelSheetOut);

        int i = 0;
        foreach (ExcelBorderExport borderExport in excelStyles.ListBorders)
        {
            string rowIdx = (i + 2).ToString();

            excelProcessor.SetCellValue(excelSheetOut, "A" + rowIdx, borderExport.SheetIndex);
            excelProcessor.SetCellValue(excelSheetOut, "B" + rowIdx, excelStyles.ListSheet.FirstOrDefault(s => s.Index == borderExport.SheetIndex).Name);
            excelProcessor.SetCellValue(excelSheetOut, "C" + rowIdx, borderExport.BorderId);

            if (borderExport.LeftBorder != null)
            {
                if (borderExport.LeftBorder.Style != null)
                    excelProcessor.SetCellValue(excelSheetOut, "D" + rowIdx, borderExport.LeftBorder.Style.ToString());

                //if (borderExport.LeftBorder.Color != null)
                //    excelProcessor.SetCellValue(excelSheetOut, "E" + rowIdx, borderExport.LeftBorder.Color.ToString());
            }

            if (borderExport.RightBorder != null)
            {
                if (borderExport.RightBorder.Style != null)
                    excelProcessor.SetCellValue(excelSheetOut, "E" + rowIdx, borderExport.RightBorder.Style.ToString());
            }

            if (borderExport.TopBorder != null)
            {
                if (borderExport.TopBorder.Style != null)
                    excelProcessor.SetCellValue(excelSheetOut, "F" + rowIdx, borderExport.TopBorder.Style.ToString());
            }

            if (borderExport.BottomBorder != null)
            {
                if (borderExport.BottomBorder.Style != null)
                    excelProcessor.SetCellValue(excelSheetOut, "G" + rowIdx, borderExport.BottomBorder.Style.ToString());
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
            proc.SetCellValue(excelSheet, "C1", "BorderId");
            proc.SetCellValue(excelSheet, "D1", "Left.Style");
            //proc.SetCellValue(excelSheet, "E1", "Left.Color");
            proc.SetCellValue(excelSheet, "E1", "Right.Style");
            proc.SetCellValue(excelSheet, "F1", "Top.Style");
            proc.SetCellValue(excelSheet, "G1", "Bottom.Style");
        }

    }
}