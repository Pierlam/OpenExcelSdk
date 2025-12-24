using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcelSdk.System;

/// <summary>
/// Represents an excel file.
/// OexExcelFile
/// </summary>
public class ExcelFile
{
    public ExcelFile(string filename, SpreadsheetDocument spreadsheetDocument)
    {
        Filename = filename;
        SpreadsheetDocument = spreadsheetDocument;
        WorkbookPart = spreadsheetDocument?.WorkbookPart;
    }

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
