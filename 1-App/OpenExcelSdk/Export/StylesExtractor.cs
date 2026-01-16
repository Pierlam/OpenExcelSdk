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
        ExcelAllStylesExport excelStyles = new ExcelAllStylesExport();


        for(int i=0; i< _proc.GetSheetCount(excelFileIn);i++)
        {
            ExcelSheet excelSheetIn = _proc.GetSheetAt(excelFileIn, i);
            ExcelSheetExport sheetTable = new ExcelSheetExport(i, excelSheetIn.Name);
            excelStyles.ListSheet.Add(sheetTable);

            // export Fills
            ExportFills(excelSheetIn, excelStyles);

            // export Borders
            ExportBorders(excelSheetIn, excelStyles);

            // export fonts
            ExportFonts(excelSheetIn, excelStyles);

            // export all styles/CellFormats
            ExportStylesSheet(excelSheetIn, excelStyles);
        }

        _proc.CloseExcelFile(excelFileIn);

        return excelStyles;
    }

    /// <summary>
    /// Export Fills.
    /// </summary>
    /// <param name="excelSheetIn"></param>
    /// <param name="excelStyles"></param>
    void ExportFills(ExcelSheet excelSheetIn, ExcelAllStylesExport excelStyles)
    {
        var stylesPart = excelSheetIn.ExcelFile.WorkbookPart.WorkbookStylesPart;

        for (int i = 0; i < stylesPart.Stylesheet.Fills.Elements().Count(); i++)
        {
            Fill fill = (Fill)excelSheetIn.ExcelFile.WorkbookPart.WorkbookStylesPart.Stylesheet.Fills.ElementAt(i);

            ExcelColor bgColor = ColorMgr.GetCellBackgroundColor(_styleMgr, excelSheetIn, fill);
            // bg color is set, not expected
            //if (bgColor != null) continue;

            ExcelColor fgColor = ColorMgr.GetCellForegroundColor(_styleMgr, excelSheetIn, fill);
            //if (fgColor == null) continue;


            //-XXDEBUG:
            if (fill.GradientFill != null)
            {
                // DEBUG: ExcelGradientFill
                int nbGradient = fill.GradientFill.Elements().Count();
                //GradientFill gradientFill = new GradientFill();
                //gradientFill.Degree =;
                //gradientFill.de
            }

            ExcelFillExport fillExport = new ExcelFillExport(excelSheetIn.Index, i, fill.PatternFill.PatternType, bgColor, fgColor);
            excelStyles.ListFills.Add(fillExport);
        }

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

            borderExport.LeftBorder = border.LeftBorder;
            borderExport.RightBorder = border.RightBorder;
            borderExport.TopBorder= border.TopBorder;
            borderExport.BottomBorder= border.BottomBorder;
            borderExport.DiagonalBorder= border.DiagonalBorder;

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
        excelStyles.ListStyle.Add(styleExport);

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
