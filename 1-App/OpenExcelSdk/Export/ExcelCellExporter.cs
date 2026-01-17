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
    static ExcelAllStylesExport _excelAllStyles;

    /// <summary>
    /// Export fist 1000 cells
    /// </summary>
    /// <param name="excelProcessor"></param>
    /// <param name="excelStyles"></param>
    /// <param name="excelFileOut"></param>
    public static void Export(ExcelProcessor excelProcessor, ExcelAllStylesExport excelAllStyles, ExcelFile excelFileIn, ExcelFile excelFileOut)
    {
        _excelAllStyles= excelAllStyles;

        ExcelSheet excelSheetOut = excelProcessor.CreateSheet(excelFileOut, "Cells");

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
        // get last row index, one-based
        int lastRowIdx = excelProcessor.GetLastRowIndex(excelSheetIn);
        int cellTotalCount = 0;
        int cellCount = 0;
        int rowNumOut = 0;
        for (int i = 1; i <= lastRowIdx; i++)
        {
            // get cells of the row
            List<ExcelCell> listCell=excelProcessor.GetRowCells(excelSheetIn, i);
            if (listCell.Count == 0) continue;

            rowNumOut= ExportCells(excelProcessor, excelSheetIn.Index, excelSheetIn.Name, i, rowNumOut, listCell, excelSheetOut, out cellCount);
            cellTotalCount += cellCount;
            if(cellCount > _excelAllStyles.CellsSheetMaxLoadCount) break;
            if (cellTotalCount > _excelAllStyles.CellsMaxLoadCount) break;
        }
    }

    /// <summary>
    /// Export the cells of the row.
    /// </summary>
    /// <param name="excelProcessor"></param>
    /// <param name="sheetIdx"></param>
    /// <param name="sheetName"></param>
    /// <param name="rowNumIn"></param>
    /// <param name="rowNumOut"></param>
    /// <param name="listCell"></param>
    /// <param name="excelSheetOut"></param>
    /// <param name="cellCount"></param>
    /// <returns></returns>
    static int ExportCells(ExcelProcessor excelProcessor, int sheetIdx, string sheetName, int rowNumIn, int rowNumOut, List<ExcelCell> listCell, ExcelSheet excelSheetOut, out int cellCount)
    {
        cellCount = listCell.Count;

        for (int i = 0; i < listCell.Count; i++)
        {
            if (i> _excelAllStyles.CellsSheetMaxLoadCount) break;
            if (i > _excelAllStyles.CellsMaxLoadCount) 
                break;

            string rowIdx = (i + rowNumOut+ 2).ToString();
            excelProcessor.SetCellValue(excelSheetOut, "A" + rowIdx, sheetIdx);
            excelProcessor.SetCellValue(excelSheetOut, "B" + rowIdx, sheetName);
            excelProcessor.SetCellValue(excelSheetOut, "C" + rowIdx, rowNumIn);
            excelProcessor.SetCellValue(excelSheetOut, "D" + rowIdx, listCell[i].Cell.CellReference.Value);

            // cell raw value?
            if(listCell[i].Cell.DataType != null)
            {
                excelProcessor.SetCellValue(excelSheetOut, "E" + rowIdx, listCell[i].Cell.DataType.ToString());
            }

            if (listCell[i].Cell.InnerText!= null)
            {
                if(int.TryParse(listCell[i].Cell.InnerText, out int intValue))
                    excelProcessor.SetCellValue(excelSheetOut, "F" + rowIdx, intValue);
                else
                    excelProcessor.SetCellValue(excelSheetOut, "F" + rowIdx, listCell[i].Cell.InnerText);
            }

            int styleIndex = 0;
            if (listCell[i].Cell.StyleIndex != null)
            {
                styleIndex = (int)listCell[i].Cell.StyleIndex.Value;
                excelProcessor.SetCellValue(excelSheetOut, "G" + rowIdx, styleIndex);

                var styleExport = _excelAllStyles.ListStyles.FirstOrDefault(s => s.StyleIndex == styleIndex);

                if(styleExport.NumberFormatId!= 0)
                    excelProcessor.SetCellValue(excelSheetOut, "H" + rowIdx, styleExport.NumberFormatId);

                excelProcessor.SetCellValue(excelSheetOut, "I" + rowIdx, styleExport.NumberFormat);

                if (styleExport.FillId != 0)
                {
                    excelProcessor.SetCellValue(excelSheetOut, "J" + rowIdx, styleExport.FillId);
                    var exportFill=  _excelAllStyles.ListFills.FirstOrDefault(f => f.FillId == styleExport.FillId);
                    if(exportFill.FgColor!=null)
                        excelProcessor.SetCellValue(excelSheetOut, "K" + rowIdx, exportFill.FgColor.Rgb);
                }

                if (styleExport.BorderId!= 0)
                {
                    excelProcessor.SetCellValue(excelSheetOut, "L" + rowIdx, styleExport.BorderId);
                }

                if (styleExport.FontId != 0)
                {
                    excelProcessor.SetCellValue(excelSheetOut, "M" + rowIdx, styleExport.FontId);
                }
            }
        }

        // nb rows used to export infos
        return listCell.Count+rowNumOut;
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
        proc.SetCellValue(excelSheet, "C1", "RowIdx");

        proc.SetCellValue(excelSheet, "D1", "CellRef");
        proc.SetCellValue(excelSheet, "E1", "DataType");
        proc.SetCellValue(excelSheet, "F1", "InnerText");

        proc.SetCellValue(excelSheet, "G1", "StyleIndex");
        proc.SetCellValue(excelSheet, "H1", "NumberFormatId");
        proc.SetCellValue(excelSheet, "I1", "NumberFormat");
        proc.SetCellValue(excelSheet, "J1", "FillId");
        proc.SetCellValue(excelSheet, "K1", "Fill.FgColor");
        proc.SetCellValue(excelSheet, "L1", "BorderId");
        proc.SetCellValue(excelSheet, "M1", "FontId");
    }

}
