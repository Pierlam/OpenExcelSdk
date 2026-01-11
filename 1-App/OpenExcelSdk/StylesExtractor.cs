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

namespace OpenExcelSdk;

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
    public ExcelStyles Extract(string filenameIn)
    {
        ExcelFile excelFileIn= _proc.OpenExcelFile(filenameIn);

        ExcelStyles excelStyles = new ExcelStyles();


        for(int i=0; i< _proc.GetSheetCount(excelFileIn);i++)
        {
            ExcelSheet excelSheetIn = _proc.GetSheetAt(excelFileIn, i);
            SheetTable sheetTable = new SheetTable(i, excelSheetIn.Name);
            excelStyles.ListSheetTable.Add(sheetTable);
            ExportStylesSheet(excelSheetIn, excelStyles);
        }

        _proc.CloseExcelFile(excelFileIn);

        return excelStyles;
    }

    void ExportStylesSheet(ExcelSheet excelSheetIn, ExcelStyles excelStyles)
    {
        var stylesPart = excelSheetIn.ExcelFile.WorkbookPart.WorkbookStylesPart;

        for (int i = 0; i < stylesPart.Stylesheet.CellFormats.Elements().Count(); i++)
        {
            ExportStyle(excelSheetIn, excelStyles, stylesPart, i);
        }
    }

    void ExportStyle(ExcelSheet excelSheetIn, ExcelStyles excelStyles, WorkbookStylesPart stylesPart, int i)
    {

        string rowIdx = (i + 2).ToString();

        // get the style of cell
        CellFormat cellFormat = (CellFormat)stylesPart.Stylesheet.CellFormats.ElementAt((int)i);

        string numberFormat=string.Empty;
        if (cellFormat.NumberFormatId != null)
        {
            numberFormat = GetNumberFormat(stylesPart, (int)cellFormat.NumberFormatId.Value);
            
        }

        StyleTable styleTable = new StyleTable(excelSheetIn.Index, i, (int)cellFormat.NumberFormatId.Value, numberFormat);
        excelStyles.ListStyleTable.Add(styleTable);

        AddStyleFill(excelSheetIn, styleTable, stylesPart, cellFormat, rowIdx);

        //if (cellFormat.ApplyAlignment == null && cellFormat.ApplyBorder == null && cellFormat.ApplyFill == null &&
        //    cellFormat.ApplyFont == null && cellFormat.ApplyProtection == null) return true;
    }


    void AddStyleFill(ExcelSheet excelSheetIn, StyleTable styleTable, WorkbookStylesPart stylesPart, CellFormat cellFormat, string rowIdx)
    {
        if (cellFormat.ApplyFill == null) return;

        Fill fill = (Fill)stylesPart.Stylesheet.Fills.ElementAt((int)cellFormat.FillId.Value);
        styleTable.FillId = (int)cellFormat.FillId.Value;
        styleTable.FillPattern = fill.PatternFill.PatternType;


        //-XXDEBUG:
        if (fill.GradientFill != null)
        {
            // DEBUG: ExcelGradientFill
            int nbGradient = fill.GradientFill.Elements().Count();
            //GradientFill gradientFill = new GradientFill();
            //gradientFill.Degree =;
            //gradientFill.de
        }


        styleTable.BgColor = ColorMgr.GetCellBackgroundColor(_styleMgr, excelSheetIn, fill);
        styleTable.FgColor = ColorMgr.GetCellForegroundColor(_styleMgr, excelSheetIn, fill);
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
