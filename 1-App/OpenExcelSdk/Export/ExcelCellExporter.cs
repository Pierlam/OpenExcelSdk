using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using OpenExcelSdk.System.Export;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcelSdk.Export;

public class ExcelCellExporter
{
    /// <summary>
    /// Export fist 1000 cells
    /// </summary>
    /// <param name="excelProcessor"></param>
    /// <param name="excelStyles"></param>
    /// <param name="excelFileOut"></param>
    public static void Export(ExcelProcessor excelProcessor, ExcelFile excelFileIn, ExcelFile excelFileOut)
    {
        // the first sheet exists already
        ExcelSheet excelSheetOut = excelProcessor.GetSheetAt(excelFileOut, 0);

        // create the out header
        CreateOutHeader(excelProcessor, excelSheetOut);

        // scan each sheet
        int nbsheet= excelProcessor.GetSheetCount(excelFileIn);

        for (int i = 0; i < nbsheet; i++)
        {
            ExcelSheet excelSheet = excelProcessor.GetSheetAt(excelFileIn, i);
            ExportSheetContent(excelProcessor, excelSheet, excelSheetOut);
        }    
    }

    static void ExportSheetContent(ExcelProcessor excelProcessor, ExcelSheet excelSheetIn, ExcelSheet excelSheetOut)
    {
        int nbRow= excelProcessor.GetLastRowIndex(excelSheetIn);
        int cellCount = 0;
        for (int i = 0; i < nbRow; i++)
        {
            ExcelRow excelRow=  excelProcessor.GetRowAt(excelSheetIn, i);

            // get cells of the row
            List<ExcelCell> listCell=excelProcessor.GetRowCells(excelRow);
            
            ExportCells(excelProcessor, excelSheetIn.Index, excelSheetIn.Name, excelRow, listCell, excelSheetOut);
        }
    }

    static void ExportCells(ExcelProcessor excelProcessor, int sheetIdx, string sheetName, ExcelRow excelRow, List<ExcelCell> listCell, ExcelSheet excelSheetOut)
    {
        for (int i = 0; i < listCell.Count; i++)
        {
            string rowIdx = (i + 2).ToString();
            excelProcessor.SetCellValue(excelSheetOut, "A" + rowIdx, sheetIdx);
            excelProcessor.SetCellValue(excelSheetOut, "B" + rowIdx, sheetName);
            excelProcessor.SetCellValue(excelSheetOut, "C" + rowIdx, listCell[i].Cell.CellReference);

            int styleIndex= 0;
            if (listCell[i].Cell.StyleIndex != null)
                styleIndex = (int)listCell[i].Cell.StyleIndex.Value;
            
            excelProcessor.SetCellValue(excelSheetOut, "D" + rowIdx, styleIndex);

            //int numberFormatId = 0;
            //if (listCell[i].Cell.Number != null)
            //    styleIndex = (int)listCell[i].Cell.StyleIndex.Value;
            //excelProcessor.SetCellValue(excelSheetOut, "D" + rowIdx, styleIndex);

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
        proc.SetCellValue(excelSheet, "C1", "Cell");
        proc.SetCellValue(excelSheet, "D1", "StyleIndex");
        proc.SetCellValue(excelSheet, "D1", "NumberFormatId");
        proc.SetCellValue(excelSheet, "E1", "NumberFormat");
        proc.SetCellValue(excelSheet, "F1", "FillId");
        proc.SetCellValue(excelSheet, "G1", "Fill.FgColor");
        proc.SetCellValue(excelSheet, "I1", "BorderId");
        proc.SetCellValue(excelSheet, "J1", "FontId");
    }

}
