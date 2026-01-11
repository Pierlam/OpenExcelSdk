using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcelSdk.System.Export;

public class StyleTable
{
    public StyleTable(int sheetIndex, int styleIdx, int numberFormatId, string numberFormat)
    {
        SheetIndex= sheetIndex;
        StyleIndex= styleIdx;
        NumberFormatId= numberFormatId;
        NumberFormat = numberFormat;
    }

    public int SheetIndex { get; set; }
    public int StyleIndex { get; set; }

    public int NumberFormatId { get; set; }

    public string NumberFormat { get; set; }

    public int FillId { get; set; }

    public string FillPattern { get; set; }

    public ExcelColor? BgColor { get; set; } = null;
    public ExcelColor? FgColor { get; set; } = null;

}
