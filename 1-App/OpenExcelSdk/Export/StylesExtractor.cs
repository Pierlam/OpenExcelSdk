using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Office2010.Word;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenExcelSdk.System;
using OpenExcelSdk.System.Export;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Fill = DocumentFormat.OpenXml.Spreadsheet.Fill;
using NumberingFormat = DocumentFormat.OpenXml.Spreadsheet.NumberingFormat;

namespace OpenExcelSdk.Export;

/// <summary>
/// Excel styles exporter.
/// </summary>
public class StylesExtractor
{
    ExcelProcessor _proc;
    StyleMgr _styleMgr;

    public StylesExtractor(ExcelProcessor excelProcessor, StyleMgr styleMgr)
    { 
        _proc = excelProcessor;
        _styleMgr= styleMgr;
    }

    /// <summary>
    /// Export the styles of the excel file into an outpiut excel file.
    /// </summary>
    /// <param name="filenameIn"></param>
    /// <param name="filenameOut"></param>
    /// <returns></returns>
    public ExcelAllStylesExport Extract(ExcelFile excelFileIn)
    {
        ExcelAllStylesExport excelAllStyles = new ExcelAllStylesExport();


        // export sharedStrings
        ExportSharedStrings(excelFileIn, excelAllStyles);

        int cellCount = 0;

        for (int i=0; i< _proc.GetSheetCount(excelFileIn);i++)
        {
            ExcelSheet excelSheetIn = _proc.GetSheetAt(excelFileIn, i);
            ExcelSheetExport sheetTable = new ExcelSheetExport(i, excelSheetIn.Name);
            excelAllStyles.ListSheets.Add(sheetTable);

            // update cells counters
            cellCount+= _proc.GetCellsCount(excelSheetIn);

            // export Fills
            if(!ExportFills(excelSheetIn, excelAllStyles))
            {
                continue;
            }

            // export Borders
            ExportBorders(excelSheetIn, excelAllStyles);

            // export fonts
            ExportFonts(excelSheetIn, excelAllStyles);

            // export all styles/CellFormats
            ExportStylesSheet(excelSheetIn, excelAllStyles);
        }

        excelAllStyles.CellsTotalCount= cellCount;

        _proc.CloseExcelFile(excelFileIn);

        return excelAllStyles;
    }

    /// <summary>
    /// Export Shared Strings.
    /// </summary>
    /// <param name="excelFileIn"></param>
    /// <param name="excelStyles"></param>
    void ExportSharedStrings(ExcelFile excelFileIn, ExcelAllStylesExport excelAllStyles)
    {
        // For shared strings, look up the value in the shared strings table.
        var stringTable = excelFileIn.WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

        if(stringTable==null)
        {
            excelAllStyles.ListError.Add("No SharedStringTablePart found in the Excel file.");
            return;
        }

        excelAllStyles.SharedStringsTotalCount = stringTable.SharedStringTable.Elements().Count();

        // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
        int i = 0;
        foreach (SharedStringItem item in stringTable.SharedStringTable.Elements<SharedStringItem>())
        {
            if(i>= excelAllStyles.SharedStringsMaxLoadCount)
            {
                break;
            }
            ExcelSharedStringExport stringExport= new ExcelSharedStringExport(i, item.InnerText);
            excelAllStyles.ListSharedStrings.Add(stringExport); 
            i++;
        }

    }

    /// <summary>
    /// Export Fills.
    /// </summary>
    /// <param name="excelSheetIn"></param>
    /// <param name="excelStyles"></param>
    bool ExportFills(ExcelSheet excelSheetIn, ExcelAllStylesExport excelStyles)
    {
        var stylesPart = excelSheetIn.ExcelFile.WorkbookPart.WorkbookStylesPart;
        if(stylesPart==null)
        {
            excelStyles.ListError.Add($"No WorkbookStylesPart found in the Excel file for sheet {excelSheetIn.Name}.");
            return false;
        }

        for (int i = 0; i < stylesPart.Stylesheet.Fills.Elements().Count(); i++)
        {
            DocumentFormat.OpenXml.Spreadsheet.Fill fill = (Fill)excelSheetIn.ExcelFile.WorkbookPart.WorkbookStylesPart.Stylesheet.Fills.ElementAt(i);

            ExcelColor fgColor = ColorMgr.GetCellForegroundColor(_styleMgr, excelSheetIn, fill);

            ExcelColor bgColor = ColorMgr.GetCellBackgroundColor(_styleMgr, excelSheetIn, fill);

            ExcelFillExport fillExport = new ExcelFillExport(excelSheetIn.Index, i, fill.PatternFill.PatternType, bgColor, fgColor);

            if (fill.GradientFill != null)
            {
                // DEBUG: ExcelGradientFill
                int nbGradient = fill.GradientFill.Elements().Count();
                foreach (DocumentFormat.OpenXml.Spreadsheet.GradientFill grad in fill.GradientFill.Elements())
                {
                    ExcelGradientFill excelGradientFill = new ExcelGradientFill(grad);
                    fillExport.ListGradient.Add(excelGradientFill);
                }
            }

            excelStyles.ListFills.Add(fillExport);
        }
        return true;
    }

    /// <summary>
    /// Export Border
    /// </summary>
    /// <param name="excelSheetIn"></param>
    /// <param name="excelStyles"></param>
    void ExportBorders(ExcelSheet excelSheetIn, ExcelAllStylesExport excelStyles)
    {
        var stylesPart = excelSheetIn.ExcelFile.WorkbookPart.WorkbookStylesPart;

        for (int i = 0; i < stylesPart.Stylesheet.Borders.Elements().Count(); i++)
        {
            Border border = (Border)excelSheetIn.ExcelFile.WorkbookPart.WorkbookStylesPart.Stylesheet.Borders.ElementAt(i);

            ExcelBorderExport borderExport = new ExcelBorderExport(excelSheetIn.Index, i);

            borderExport.LeftBorder = ExcelCellBorderUtils.CreateCellBorder(border.LeftBorder);
            borderExport.RightBorder = ExcelCellBorderUtils.CreateCellBorder(border.RightBorder);
            borderExport.TopBorder= ExcelCellBorderUtils.CreateCellBorder(border.TopBorder);
            borderExport.BottomBorder= ExcelCellBorderUtils.CreateCellBorder(border.BottomBorder);
            borderExport.DiagonalBorder= ExcelCellBorderUtils.CreateCellBorder(border.DiagonalBorder);

            excelStyles.ListBorders.Add(borderExport);

        }
    }

    /// <summary>
    /// Export all defined fonts.
    /// </summary>
    /// <param name="excelSheetIn"></param>
    /// <param name="excelStyles"></param>
    void ExportFonts(ExcelSheet excelSheetIn, ExcelAllStylesExport excelStyles)
    {
        var stylesPart = excelSheetIn.ExcelFile.WorkbookPart.WorkbookStylesPart;

        for (int i = 0; i < stylesPart.Stylesheet.Fonts.Elements().Count(); i++)
        {
            Font font = (Font)excelSheetIn.ExcelFile.WorkbookPart.WorkbookStylesPart.Stylesheet.Fonts.ElementAt(i);

            ExcelFontExport fontExport = new ExcelFontExport(excelSheetIn.Index, i);

            excelStyles.ListFonts.Add(fontExport);
        }
    }

    /// <summary>
    /// Export Styles, i.e. CellFormats.
    /// </summary>
    /// <param name="excelSheetIn"></param>
    /// <param name="excelStyles"></param>
    void ExportStylesSheet(ExcelSheet excelSheetIn, ExcelAllStylesExport excelStyles)
    {
        var stylesPart = excelSheetIn.ExcelFile.WorkbookPart.WorkbookStylesPart;

        for (int i = 0; i < stylesPart.Stylesheet.CellFormats.Elements().Count(); i++)
        {
            ExportStyle(excelSheetIn, excelStyles, stylesPart, i);
        }
    }

    void ExportStyle(ExcelSheet excelSheetIn, ExcelAllStylesExport excelStyles, WorkbookStylesPart stylesPart, int i)
    {

        string rowIdx = (i + 2).ToString();

        // get the style of cell
        CellFormat cellFormat = (CellFormat)stylesPart.Stylesheet.CellFormats.ElementAt((int)i);

        string numberFormat=string.Empty;
        if (cellFormat.NumberFormatId != null)
        {
            numberFormat = GetNumberFormat(stylesPart, (int)cellFormat.NumberFormatId.Value);
            
        }

        ExcelStyleExport styleExport = new ExcelStyleExport(excelSheetIn.Index, i, (int)cellFormat.NumberFormatId.Value, numberFormat);
        excelStyles.ListStyles.Add(styleExport);

        if (cellFormat.ApplyFill != null)
        {
            Fill fill = (Fill)stylesPart.Stylesheet.Fills.ElementAt((int)cellFormat.FillId.Value);
            styleExport.FillId = (int)cellFormat.FillId.Value;
        }

        if (cellFormat.ApplyBorder != null)
        {
            Border border = (Border)stylesPart.Stylesheet.Borders.ElementAt((int)cellFormat.BorderId.Value);
            styleExport.BorderId = (int)cellFormat.BorderId.Value;
        }

        if (cellFormat.ApplyFont != null)
        {
            Font font = (Font)stylesPart.Stylesheet.Borders.ElementAt((int)cellFormat.FontId.Value);
            styleExport.FontId = (int)cellFormat.FontId.Value;
        }

        //if (cellFormat.ApplyAlignment == null && cellFormat.ApplyBorder == null && cellFormat.ApplyFill == null &&
        //    cellFormat.ApplyFont == null && cellFormat.ApplyProtection == null) return true;
    }


    /// <summary>
    /// Get the number format string from its id.
    /// works also for buil-oin format like 44:currency.
    /// </summary>
    /// <param name="stylesPart"></param>
    /// <param name="formatId"></param>
    /// <returns></returns>
    string GetNumberFormat(WorkbookStylesPart stylesPart, int formatId)
    {
        if (stylesPart.Stylesheet.NumberingFormats == null) return string.Empty;

        foreach (NumberingFormat nf in stylesPart.Stylesheet.NumberingFormats.Elements<NumberingFormat>())
        {
            if (nf.NumberFormatId.Value == formatId)
            {
                return nf.FormatCode.Value;
            }
        }

        return string.Empty;
    }
}
