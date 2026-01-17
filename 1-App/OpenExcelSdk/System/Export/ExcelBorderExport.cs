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
    public ExcelCellBorder? LeftBorder { get; set; } = null;

    public ExcelCellBorder? RightBorder { get; set; } = null;

    public ExcelCellBorder? TopBorder { get; set; } = null;

    public ExcelCellBorder? BottomBorder { get; set; } = null;
    public ExcelCellBorder? DiagonalBorder { get; set; } = null;

}
