using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcelSdk.System.Export;

/// <summary>
/// One style (which is CellFormat) entry in the style table.
/// Style is identified by its StyleIndex in the Styles part.
/// </summary>
public class ExcelStyleExport
{
    public ExcelStyleExport(int sheetIndex, int styleIdx, int numberFormatId, string numberFormat)
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

    public int BorderId { get; set; }

    public int FontId { get; set; }

}
