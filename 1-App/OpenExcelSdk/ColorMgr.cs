using DocumentFormat.OpenXml.Spreadsheet;
using OpenExcelSdk.System;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcelSdk;

public class ColorMgr
{
    /// <summary>
    /// Set a foreground color to a cell, with pattern set to Solid.
    /// So the background color has no effect.
    /// If the cell has alreayd this foreground color, return the cell color without any modification.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="excelCell"></param>
    /// <param name="rgb"></param>
    /// <returns></returns>
    public static ExcelCellColor SetCellFgColor(StyleMgr styleMgr, ExcelSheet excelSheet, ExcelCell excelCell, string rgb)
    {
        if (string.IsNullOrEmpty(rgb)) return null;
        if (!rgb.StartsWith("#")) return null;
        if (rgb.Length!=7) return null;

        ExcelCellColor excelCellColor=  GetCellColor(styleMgr, excelSheet, excelCell);

        if (excelCellColor != null && excelCellColor.FgColor!=null)
        {
            // The cell has already this color ?
            if (excelCellColor.FgColor.Rgb.Equals(rgb) && excelCellColor.BgColor==null) return excelCellColor;
        }

        // is there a fill with the same fg color (+bg=null) ?
        Fill fillFound = GetFillByFgColorRgb(styleMgr, excelSheet, rgb, out int indexFillFound);

        var stylesPart = excelSheet.ExcelFile.WorkbookPart.WorkbookStylesPart;

        CellFormat cellFormat = ExcelUtils.GetCellFormat(excelSheet, excelCell);
        if (fillFound != null)
        {
            // is there a CellFormat matching ?
            cellFormat = styleMgr.FindCellFormatWithFgColor(stylesPart, cellFormat, indexFillFound, out int index);
            if (cellFormat != null)
            {
                excelCell.Cell.StyleIndex = (uint)index;
                excelCell.CellFormat = cellFormat;
                return GetCellColor(styleMgr, excelSheet, excelCell);
            }
        }
        else
        {
            indexFillFound = styleMgr.CreateFill(excelSheet, rgb);
        }


        // have to create a new new CellFormat, copy other parts
        styleMgr.CreateCellFormatSetFillId(excelSheet, excelCell, indexFillFound, out int styleIndex);

        //excelCell.Cell.StyleIndex = (uint)styleIndex;

        return GetCellColor(styleMgr, excelSheet, excelCell);
    }

    /// <summary>
    /// Is there a fill with the same fg color (+bg=null) ?
    /// </summary>
    /// <param name="rgb"></param>
    /// <returns></returns>
    static Fill GetFillByFgColorRgb(StyleMgr styleMgr, ExcelSheet excelSheet, string rgb, out int indexFillFound)
    {
        indexFillFound = -1;

        for (int i = 0; i < excelSheet.ExcelFile.WorkbookPart.WorkbookStylesPart.Stylesheet.Fills.Elements().Count(); i++)
        {
            Fill fill = (Fill)excelSheet.ExcelFile.WorkbookPart.WorkbookStylesPart.Stylesheet.Fills.ElementAt(i);

            ExcelColor bgColor = GetCellBackgroundColor(styleMgr, excelSheet, fill);
            // bg color is set, not expected
            if (bgColor != null) continue;

            ExcelColor fgColor = GetCellForegroundColor(styleMgr, excelSheet, fill);
            if (fgColor == null) continue;

            indexFillFound = i;
            if (fgColor.Rgb.Equals(rgb)) return fill;
        }

        // not found
        return null;
    }


    /// <summary>
    /// Get the foreground and background color of the cell.
    /// background and/or foreground color can be null if there is no color defined.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="excelCell"></param>
    /// <returns></returns>
    public static ExcelCellColor GetCellColor(StyleMgr styleMgr, ExcelSheet excelSheet, ExcelCell excelCell)
    {
        CellFormat cellFormat = ExcelUtils.GetCellFormat(excelSheet, excelCell);
        if (cellFormat == null) return null;

        Fill fill = ColorMgr.GetCellFill(excelSheet, excelCell);

        ExcelColor bgColor = GetCellBackgroundColor(styleMgr, excelSheet, fill);
        ExcelColor fgColor = GetCellForegroundColor(styleMgr, excelSheet, fill);
        ExcelCellColor excelCellColor = new ExcelCellColor(excelCell, cellFormat.FillId, bgColor, fgColor);
                
        //-XXDEBUG:
        //if(fill.GradientFill!=null)
        //{
        //    // DEBUG:
        //    int nbGradient = fill.GradientFill.Elements().Count();
        //    GradientFill gradientFill = new GradientFill();
        //    //gradientFill.Degree
        //}

        if(bgColor!=null)
            bgColor.ColorName = GetColorName(bgColor.Rgb);
        if (fgColor != null)
            fgColor.ColorName = GetColorName(fgColor.Rgb);

        return excelCellColor;
    }

    public static ExcelColor GetCellBackgroundColor(StyleMgr styleMgr,  ExcelSheet excelSheet, Fill fill)
    {
        if (fill.PatternFill.BackgroundColor == null) return null;

        BackgroundColor bgColor = fill.PatternFill.BackgroundColor;


        // 2 cases: direct color or theme color
        if (bgColor.Rgb != null)
        {
            ExcelColor color = new ExcelColor(bgColor.Rgb.ToString());
            return color;
        }

        if (bgColor.Theme != null)
        {
            int themeIndex = (int)bgColor.Theme.Value;
            double tint = bgColor.Tint != null ? bgColor.Tint.Value : 0;

            string rgb = styleMgr.GetThemeColor(excelSheet.ExcelFile.WorkbookPart, themeIndex, tint);
            ExcelColor color = new ExcelColor(themeIndex, rgb);
            return color;
        }
        return null;
    }

    public static ExcelColor GetCellForegroundColor(StyleMgr styleMgr, ExcelSheet excelSheet, Fill fill)
    {
        if (fill.PatternFill.ForegroundColor == null) return null;
       
        ForegroundColor fgColor = fill.PatternFill.ForegroundColor;

        // 2 cases: direct color or theme color
        if (fgColor.Rgb != null)
        {
            // it's a argb value in fact, A+RGB
            ExcelColor color = new ExcelColor(fgColor.Rgb.ToString());
            return color;
        }

        if (fgColor.Theme != null)
        {
            int themeIndex = (int)fgColor.Theme.Value;
            double tint = fgColor.Tint != null ? fgColor.Tint.Value : 0;

            // std RGB value
            string rgb = styleMgr.GetThemeColor(excelSheet.ExcelFile.WorkbookPart, themeIndex, tint);
            ExcelColor color = new ExcelColor(themeIndex, rgb);
            return color;
        }
        return null;
    }

    /// <summary>
    /// Return the fill (bg and fg colors) object of the cell.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="excelCell"></param>
    /// <param name="rgb"></param>
    /// <returns></returns>
    public static Fill GetCellFill(ExcelSheet excelSheet, ExcelCell excelCell)
    {
        CellFormat cellFormat = ExcelUtils.GetCellFormat(excelSheet, excelCell);
        if (cellFormat == null) return null;

        if (cellFormat.FillId == null) return null;
        
        uint fillId = cellFormat.FillId.Value;
        return excelSheet.ExcelFile.WorkbookPart.WorkbookStylesPart.Stylesheet.Fills.ElementAt((int)fillId) as Fill;        
    }

    /// <summary>
    /// Get the color base on the hexa code if it exists.
    /// </summary>
    /// <param name="rgb"></param>
    /// <returns></returns>
    static ColorName GetColorName(string rgb)
    {
        if(string.IsNullOrWhiteSpace(rgb))return ColorName.Undefined;
        if (rgb.Equals("#FFFF00")) return ColorName.Yellow;
        if (rgb.Equals("#FF0000")) return ColorName.Red;
        if (rgb.Equals("#0000FF")) return ColorName.Blue;

        return ColorName.Undefined;
    }

}
