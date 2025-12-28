using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenExcelSdk;

public class ExcelRow
{
    public ExcelRow(Row row)
    {
        Row = row;
    }

    /// <summary>
    /// Open Xml row object.
    /// </summary>
    public Row Row { get; set; }
}