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

    /// <summary>
    /// Get the row index, not the address.
    /// </summary>
    /// <returns></returns>
    public int GetRowIndex()
    {
        return (int)Row.RowIndex.Value;
    }
}