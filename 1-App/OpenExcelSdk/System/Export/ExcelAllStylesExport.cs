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
    /// Total count of shared strings in the excel file.
    /// </summary>
    public int SharedStringsMaxLoadCount { get; set; } = 1000;

    /// <summary>
    /// Cells max load total, all sheets.
    /// </summary>
    public int CellsMaxLoadCount { get; set; } = 5000;

    /// <summary>
    /// Cells max load by sheet.
    /// </summary>
    public int CellsSheetMaxLoadCount { get; set; } = 1000;

    /// <summary>
    /// total count of shared strings in the excel file.
    /// </summary>
    public int SharedStringsTotalCount { get; set; } = 0;

    /// <summary>
    /// Total count of cells in the excel file.
    /// </summary>
    public int CellsTotalCount { get; set; } = 0;

    /// <summary>
    /// Not all shared strings are loaded.
    /// </summary>
    public List<ExcelSharedStringExport> ListSharedStrings { get; set; } = new List<ExcelSharedStringExport>();  

    /// <summary>
    /// Sheet table list.
    /// SheetExport
    /// </summary>
    public List<ExcelSheetExport> ListSheets { get; set; } = new List<ExcelSheetExport>();

    /// <summary>
    /// Style (CellFormat) table list.
    /// </summary>
    public List<ExcelStyleExport> ListStyles { get; set; } = new List<ExcelStyleExport>();

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


    /// <summary>
    /// List of errors during the extraction.
    /// </summary>
    public List<string> ListError { get; set; } = new List<string>();

    /// <summary>
    /// e.g.  No SharedStringTablePart found in the Excel file.
    /// </summary>
    public List<string> ListInfo { get; set; } = new List<string>();

}
