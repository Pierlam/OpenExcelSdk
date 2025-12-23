using DocumentFormat.OpenXml;
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
    /// <summary>
    /// last id of custom format created.
    /// min=164
    /// </summary>
    uint _customFormatIdMax = 0;

    /// <summary>
    /// Create a new style to update numberFormatId, in Cellformat.
    /// Clone the style and modify it.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="excelCell"></param>
    /// <returns></returns>
    public bool UpdateCellStyleNumberFormatId(ExcelSheet excelSheet, ExcelCell excelCell, uint? numberFormatId)
    {
        // the cell has no style
        //if (excelCell.Cell.StyleIndex == null) return true;

        var stylesPart = excelSheet.ExcelFile.WorkbookPart.WorkbookStylesPart;

        CellFormat? currCellFormat = null;
        CellFormat? cellFormat = null;
        Alignment? alignment = null;
        BooleanValue? applyAlignement = null;
        BooleanValue? applyBorder = null;
        BooleanValue? applyFill = null;
        BooleanValue? applyFont = null;
        BooleanValue? applyNumberFormat = null;
        BooleanValue? applyProtection = null;
        UInt32Value? borderId = null;
        UInt32Value? fillId = null;
        UInt32Value? fontId = null;
        UInt32Value? numFmtId = null;
        Protection? protection = null;

        // set the new value or set null?
        if (numberFormatId != null)
        {
            applyNumberFormat = true;
            numFmtId = (uint)numberFormatId;
        }

        if (excelCell.Cell.StyleIndex != null)
        {
            // get the style of cell 
            currCellFormat = (CellFormat)stylesPart.Stylesheet.CellFormats.ElementAt((int)excelCell.Cell.StyleIndex.Value);

            // same style/cellFormat already exists?
            cellFormat = FindCellFormatWithNumberFormatId(stylesPart, currCellFormat, numberFormatId, out int index);
            if (cellFormat != null)
            {
                // get the index of the found style
                excelCell.Cell.StyleIndex = (uint)index;
                return true;
            }
            alignment = currCellFormat.Alignment;
            applyAlignement= currCellFormat.ApplyAlignment;
            applyBorder= currCellFormat.ApplyBorder;
            applyFill = currCellFormat.ApplyFill;
            applyFont = currCellFormat.ApplyFont;
            applyProtection= currCellFormat.ApplyProtection;
            borderId= currCellFormat.BorderId;
            fillId= currCellFormat.FillId;
            fontId= currCellFormat.FontId;
            protection= currCellFormat.Protection;
        }

        // need to create a new style
        cellFormat = new CellFormat { 
            Alignment = alignment,
            ApplyAlignment = applyAlignement,
            ApplyBorder = applyBorder,
            ApplyFill =  applyFill,
            ApplyFont = applyFont,
            // apply a value or null
            ApplyNumberFormat = applyNumberFormat,
            ApplyProtection = applyProtection,
            BorderId = borderId,
            FillId = fillId,
            FontId = fontId,

            // apply a value or null
            NumberFormatId = numFmtId,
            Protection = protection
        };

        // append the new style
        stylesPart.Stylesheet.CellFormats.AppendChild(cellFormat);

        stylesPart.Stylesheet.Save();

        // get the index and set to cell
        int count = stylesPart.Stylesheet.CellFormats.Elements().Count();
        excelCell.Cell.StyleIndex= (uint)(count - 1);
        return true;
    }

    /// <summary>
    /// Get the NumberFormatID (built-in or custom) or create a format (custom).
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="format"></param>
    /// <param name="formatId"></param>
    /// <param name="error"></param>
    /// <returns></returns>
    public bool GetOrCreateNumberFormat(ExcelSheet excelSheet, string format, out uint formatId, out ExcelError error)
    {
        error = null;

        // is the format a Built-In format?
        if (!BuiltInNumberFormatMgr.GetFormatId(format, out formatId))
        {
            // is the format a custom format?
            if (!GetCustomNumberFormatId(excelSheet, format, out formatId))
            {
                // create a new custom format
                if (!CreateCustomNumberFormat(excelSheet, format, out formatId, out error))
                    return false;                
            }
        }
        return true;
    }

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
        if (excelCell.Cell.StyleIndex == null) return false;

        // get the style and then the cell format
        var stylesPart = excelSheet.ExcelFile.WorkbookPart.WorkbookStylesPart;
        var cellFormat = (CellFormat)stylesPart.Stylesheet.CellFormats.ElementAt((int)excelCell.Cell.StyleIndex.Value);

        // no number format
        if (cellFormat.ApplyNumberFormat == null) return false;

        // has number format
        return true;
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
    /// same style/cellFormat already exists?
    /// </summary>
    /// <param name="cellFormat"></param>
    /// <param name="numberFormatId"></param>
    /// <returns></returns>
    public CellFormat? FindCellFormatWithNumberFormatId(WorkbookStylesPart stylesPart, CellFormat cellFormat, uint? numberFormatId, out int index)
    {
        index = 0;
        if (cellFormat==null) return null;
        if(stylesPart.Stylesheet==null) return null;

        UInt32Value? numFmtId = null; 
        if(numberFormatId!=null) numFmtId= (uint)numberFormatId;

        // scan each cell format
        for (int i=0; i<stylesPart.Stylesheet.CellFormats.Elements().Count(); i++)
        {
            index = i;
            var cellFormatFound = (CellFormat)stylesPart.Stylesheet.CellFormats.ElementAt(i);
            if(cellFormatFound.Alignment== cellFormat.Alignment &&
                cellFormatFound.BorderId== cellFormat.BorderId &&
                cellFormatFound.FillId == cellFormat.FillId &&
                cellFormatFound.Protection == cellFormat.Protection &&
                cellFormatFound.FontId == cellFormat.FontId &&
                cellFormatFound.NumberFormatId== numFmtId)
                return cellFormatFound;
        }
        return null;
    }

    /// <summary>
    /// Create a new custom format.
    ///
    /// IDs 0-163 are reserved by Excel
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="format"></param>
    /// <param name="formatId"></param>
    /// <param name="error"></param>
    /// <returns></returns>
    public bool CreateCustomNumberFormat(ExcelSheet excelSheet, string format, out uint formatId, out ExcelError error)
    {
        error = null;

        var stylesPart = excelSheet.ExcelFile.WorkbookPart.WorkbookStylesPart;

        // max id not yet calculated?
        if (_customFormatIdMax == 0)
        {
            // Get the maximum numFmtId from all NumberingFormat elements
            _customFormatIdMax = stylesPart.Stylesheet.NumberingFormats
                .Elements<NumberingFormat>()
                .Select(nf => nf.NumberFormatId.Value)
                .Max();
        }

        // use the new one
        _customFormatIdMax++;

        uint customFormatId = _customFormatIdMax;

        // create a new custom format 
        stylesPart.Stylesheet.NumberingFormats.Append(new NumberingFormat()
        {
            NumberFormatId = (uint)customFormatId,
            FormatCode = StringValue.FromString(format)
        });
        stylesPart.Stylesheet.Save();
        formatId= (uint)customFormatId;
        return true;
    }

    /// <summary>
    /// Get number format string from formatId, only custom formats.
    /// If built-in format, return false. Not saved in the Styles part!
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="formatId"></param>
    /// <param name="dataFormat"></param>
    /// <returns></returns>
    public bool GetCustomNumberFormat(ExcelSheet excelSheet, uint formatId, out string dataFormat)
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

    public bool GetCustomNumberFormatId(ExcelSheet excelSheet, string dataFormat, out uint formatId)
    {
        var stylesheet = excelSheet.ExcelFile.WorkbookPart.WorkbookStylesPart.Stylesheet;
        if (stylesheet.NumberingFormats != null)
        {
            foreach (NumberingFormat nf in stylesheet.NumberingFormats.Elements<NumberingFormat>())
            {
                if (nf.FormatCode.Value == dataFormat)
                {
                    formatId = nf.NumberFormatId.Value;
                    return true;
                }
            }
        }

        // Built-in format
        formatId = 0;
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
