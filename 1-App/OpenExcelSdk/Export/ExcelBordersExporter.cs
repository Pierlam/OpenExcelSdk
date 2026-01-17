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
            excelProcessor.SetCellValue(excelSheetOut, "B" + rowIdx, excelStyles.ListSheets.FirstOrDefault(s => s.Index == borderExport.SheetIndex).Name);
            excelProcessor.SetCellValue(excelSheetOut, "C" + rowIdx, borderExport.BorderId);

            if (borderExport.LeftBorder != null)
            {
                    excelProcessor.SetCellValue(excelSheetOut, "D" + rowIdx, borderExport.LeftBorder.BorderStyle.ToString());

                //if (borderExport.LeftBorder.Color != null)
                //    excelProcessor.SetCellValue(excelSheetOut, "E" + rowIdx, borderExport.LeftBorder.Color.ToString());
            }

            if (borderExport.RightBorder != null)
            {
                    excelProcessor.SetCellValue(excelSheetOut, "E" + rowIdx, borderExport.RightBorder.BorderStyle.ToString());
            }

            if (borderExport.TopBorder != null)
            {
                    excelProcessor.SetCellValue(excelSheetOut, "F" + rowIdx, borderExport.TopBorder.BorderStyle.ToString());
            }

            if (borderExport.BottomBorder != null)
            {
                    excelProcessor.SetCellValue(excelSheetOut, "G" + rowIdx, borderExport.BottomBorder.BorderStyle.ToString());
            }

            if (borderExport.DiagonalBorder!= null)
            {
                excelProcessor.SetCellValue(excelSheetOut, "H" + rowIdx, borderExport.DiagonalBorder.BorderStyle.ToString());
            }

            i++;
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
            proc.SetCellValue(excelSheet, "H1", "Diagonal.Style");
        }

    }
}