using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcelSdk.System;

/// <summary>
/// Excel cell color, can be foreground or background cell color.
/// </summary>
public class ExcelCellColor
{
    public ExcelCellColor(ExcelCell excelCell, uint fillId, ExcelColor BgColor, ExcelColor fgColor)
    {
        ExcelCell = excelCell;
        FillId = fillId;
        BgColor = BgColor;
        FgColor = fgColor;
    }

    public ExcelCell ExcelCell { get; set; }

    /// <summary>
    /// OpenXml Fill object id, to find the Fill object.
    /// Contains PatternFill and GradientFill objects.
    /// </summary>
    public uint FillId { get; set; }

    /// <summary>
    /// Background color, can be null.
    /// </summary>
    public ExcelColor? BgColor { get; set; }

    /// <summary>
    /// Foreground color, can be null.
    /// </summary>
    public ExcelColor? FgColor { get; set; }

}
