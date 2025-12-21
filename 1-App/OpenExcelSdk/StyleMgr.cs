using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenExcelSdk.System;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcelSdk;
public class StyleMgr
{
    public bool GetCellNumberFormatId(ExcelSheet excelSheet, ExcelCell excelCell, out uint numFmtId)
    {
        numFmtId = 0;

        //--no style
        if (excelCell.Cell.StyleIndex == null) return true;

        var stylesPart = excelSheet.ExcelFile.WorkbookPart.WorkbookStylesPart;
        var cellFormat = (CellFormat)stylesPart.Stylesheet.CellFormats.ElementAt((int)excelCell.Cell.StyleIndex.Value);

        if (cellFormat.NumberFormatId == null) return true;

        numFmtId = cellFormat.NumberFormatId.Value;
        return true;
    }

    /// <summary>
    /// Has the style cell format set to a value.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="excelCell"></param>
    /// <returns></returns>
    public bool HasCellFormat(ExcelSheet excelSheet, ExcelCell excelCell)
    {
        // no style, ok
        if (excelCell.Cell.StyleIndex == null) return true;

        // get the style and then the cell format
        var stylesPart = excelSheet.ExcelFile.WorkbookPart.WorkbookStylesPart;
        var cellFormat = (CellFormat)stylesPart.Stylesheet.CellFormats.ElementAt((int)excelCell.Cell.StyleIndex.Value);

        if (cellFormat.ApplyNumberFormat != null) return true;
        return false;
    }

    /// <summary>
    /// All others style than cell format are null, not set to a value.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="excelCell"></param>
    /// <returns></returns>
    public bool AllOthersStyleThanFormatAreNull(ExcelSheet excelSheet, ExcelCell excelCell)
    {
        // no style, ok
        if (excelCell.Cell.StyleIndex == null) return true;

        // get the style and then the cell format
        var stylesPart = excelSheet.ExcelFile.WorkbookPart.WorkbookStylesPart;
        var cellFormat = (CellFormat)stylesPart.Stylesheet.CellFormats.ElementAt((int)excelCell.Cell.StyleIndex.Value);

        // all others styles are null
        if(cellFormat.ApplyAlignment == null && cellFormat.ApplyBorder== null && cellFormat.ApplyFill == null && 
            cellFormat.ApplyFont == null && cellFormat.ApplyProtection == null) return true;

        return false;
    }

    /// <summary>
    /// Get number format string from formatId
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="formatId"></param>
    /// <param name="dataFormat"></param>
    /// <returns></returns>
    public bool GetNumberFormat(ExcelSheet excelSheet, uint formatId, out string dataFormat)
    {
        var stylesheet = excelSheet.ExcelFile.WorkbookPart.WorkbookStylesPart.Stylesheet;
        if (stylesheet.NumberingFormats != null)
        {
            foreach (NumberingFormat nf in stylesheet.NumberingFormats.Elements<NumberingFormat>())
            {
                if (nf.NumberFormatId.Value == formatId)
                {
                    dataFormat = nf.FormatCode.Value;
                    return true;
                }
            }
        }

        // Built-in format
        dataFormat=string.Empty;
        return false;
    }

    /// <summary>
    /// Remove the formula from a cell.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="excelCell"></param>
    /// <returns></returns>
    public bool RemoveFormula(ExcelSheet excelSheet, ExcelCell excelCell)
    {
        if (excelCell.Cell.CellFormula == null) return true;

        CalculationChainPart calculationChainPart = excelSheet.ExcelFile.WorkbookPart.CalculationChainPart;
        CalculationChain calculationChain = calculationChainPart.CalculationChain;
        var calculationCells = calculationChain.Elements<CalculationCell>().ToList();

        string cellRef = excelCell.Cell.CellReference;
        CalculationCell calculationCell = calculationCells.Where(c => c.CellReference == cellRef).FirstOrDefault();

        excelCell.Cell.CellFormula.Remove();
        if (calculationCell != null)
        {
            calculationCell.Remove();
            calculationCells.Remove(calculationCell);
        }
        // remove formula from the cell
        excelCell.Cell.CellFormula = null;

        // if there is no more calculation cell, remove the calculation chain part
        if (calculationCells.Count == 0)
            excelSheet.ExcelFile.WorkbookPart.DeletePart(calculationChainPart);

        return true;
    }

}
