using DocumentFormat.OpenXml.Packaging;

namespace OpenExcelSdk;

/// <summary>
/// Represents an excel file.
/// </summary>
public class ExcelFile
{
    public ExcelFile(string filename, SpreadsheetDocument spreadsheetDocument)
    {
        Filename = filename;
        SpreadsheetDocument = spreadsheetDocument;
        WorkbookPart = spreadsheetDocument?.WorkbookPart;
    }

    /// <summary>
    /// Excel filename without path and extension.
    /// </summary>
    public string Filename { get; set; }

    /// <summary>
    /// OpenXml excel object.
    /// </summary>
    public SpreadsheetDocument SpreadsheetDocument { get; set; }

    /// <summary>
    /// OpenXml Worksheet object.
    /// </summary>
    public WorkbookPart WorkbookPart { get; set; }
}