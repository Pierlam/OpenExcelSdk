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
    /// Get the foreground and background color of the cell.
    /// background and/or foreground color can be null if there is no color defined.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="excelCell"></param>
    /// <returns></returns>
    public static ExcelCellColor GetCellColor(StyleMgr styleMgr, ExcelSheet excelSheet, ExcelCell excelCell, CellFormat cellFormat)
    {
        if (cellFormat == null) return null;

        if (cellFormat.FillId == null) return null;

        uint fillId = cellFormat.FillId.Value;
        Fill fill = excelSheet.ExcelFile.WorkbookPart.WorkbookStylesPart.Stylesheet.Fills.ElementAt((int)fillId) as DocumentFormat.OpenXml.Spreadsheet.Fill;

        ExcelColor bgColor = GetCellBackgroundColor(styleMgr, excelSheet, fill);
        ExcelColor fgColor = GetCellForegroundColor(styleMgr, excelSheet, fill);
        ExcelCellColor excelCellColor = new ExcelCellColor(excelCell, fillId, bgColor, fgColor);

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

    static ColorName GetColorName(string rgb)
    {
        if(string.IsNullOrWhiteSpace(rgb))return ColorName.Undefined;
        if (rgb.Equals("FFFF00")) return ColorName.Yellow;
        if (rgb.Equals("FF0000")) return ColorName.Red;
        if (rgb.Equals("0000FF")) return ColorName.Blue;

        return ColorName.Undefined;
    }

}
