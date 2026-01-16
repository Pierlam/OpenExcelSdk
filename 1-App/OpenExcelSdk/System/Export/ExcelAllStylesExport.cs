using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcelSdk.System.Export;

/// <summary>
/// All styles of an Excel file.
/// </summary>
public class ExcelAllStylesExport
{
    /// <summary>
    /// Sheet table list.
    /// SheetExport
    /// </summary>
    public List<ExcelSheetExport> ListSheet { get; set; } = new List<ExcelSheetExport>();

    /// <summary>
    /// Style (CellFormat) table list.
    /// </summary>
    public List<ExcelStyleExport> ListStyle { get; set; } = new List<ExcelStyleExport>();

    /// <summary>
    /// Fill list.
    /// </summary>
    public List<ExcelFillExport> ListFills { get; set; } = new List<ExcelFillExport>();

    /// <summary>
    /// Border list.
    /// </summary>
    public List<ExcelBorderExport> ListBorders { get; set; } = new List<ExcelBorderExport>();

    /// <summary>
    /// Font list.
    /// </summary>
    public List<ExcelFontExport> ListFonts { get; set; } = new List<ExcelFontExport>();
}
