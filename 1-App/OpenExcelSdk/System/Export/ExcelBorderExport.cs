using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcelSdk.System.Export;

public class ExcelBorderExport
{
    public ExcelBorderExport(int sheetIndex, int borderId)
    {
        SheetIndex = sheetIndex;
        BorderId = borderId;
    }

    public int SheetIndex { get; set; }
    public int BorderId { get; set; }

    /// <summary>
    /// OpenXml Object.
    /// TODO: replace by custom object.
    /// </summary>
    public LeftBorder? LeftBorder { get; set; } = null;

    public RightBorder? RightBorder { get; set; } = null;

    public TopBorder? TopBorder { get; set; } = null;

    public BottomBorder? BottomBorder { get; set; } = null;
    public DiagonalBorder? DiagonalBorder { get; set; } = null;

}
