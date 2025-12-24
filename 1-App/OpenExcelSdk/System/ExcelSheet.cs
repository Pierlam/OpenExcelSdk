using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcelSdk.System;

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

    public ExcelFile ExcelFile { get; set; }

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
