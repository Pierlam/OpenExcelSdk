using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenExcelSdk;

/// <summary>
/// An Excel sheet.
/// </summary>
public class ExcelSheet
{
    public ExcelSheet(ExcelFile excelFile, Sheet sheet)
    {
        ExcelFile = excelFile;
        Sheet = sheet;
        WorksheetPart worksheetPart = (WorksheetPart)ExcelFile.WorkbookPart.GetPartById(Sheet.Id);
        Worksheet = worksheetPart.Worksheet;
        Rows = worksheetPart.Worksheet.Descendants<Row>();
    }

    /// <summary>
    /// The excel file hsoting the sheet.
    /// </summary>
    public ExcelFile ExcelFile { get; set; }

    /// <summary>
    /// Index of the sheet, first one index is 0.
    /// </summary>
    public int Index { get; set; }

    /// <summary>
    /// Name of the sheet.
    /// </summary>
    public string Name { get; set; }

    /// <summary>
    /// OpenXml Sheet object.
    /// </summary>
    public Sheet Sheet { get; set; }

    /// <summary>
    /// OpenXml Worksheet object.
    /// </summary>
    public Worksheet Worksheet { get; set; }

    /// <summary>
    /// OpenXml Rows object.
    /// </summary>
    public IEnumerable<Row> Rows { get; set; }
}