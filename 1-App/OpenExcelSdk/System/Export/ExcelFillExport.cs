using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcelSdk.System.Export;

/// <summary>
/// Excel sheet fill.
/// </summary>
public class ExcelFillExport
{
    public ExcelFillExport(int sheetIndex, int fillId, string patternType, ExcelColor? bgColor, ExcelColor? fgColor)
    {
        SheetIndex= sheetIndex;
        FillId= fillId;
        PatternType= patternType;
        BgColor= bgColor;
        FgColor= fgColor;
    }

    public int SheetIndex { get; set; }
    public int FillId { get; set; }
    public string PatternType { get; set; }
    public ExcelColor? BgColor { get; set; } = null;
    public ExcelColor? FgColor { get; set; } = null;

    /// <summary>
    /// Not yet managed.
    ///     fill.GradientFill.Elements().Count();
    /// </summary>
    public List<ExcelGradientFill> ListGradient { get; set; } = new List<ExcelGradientFill>();
    
}
