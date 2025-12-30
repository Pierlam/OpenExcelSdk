namespace OpenExcelSdk;

/// <summary>
/// Excel cell value, can manage all cases of type/value.
/// Type/Value can be: stirng, integer, double, dateOnly,, dateTime and timeOnly.
/// Contains also number format used to format the display of the value.
/// </summary>
public class ExcelCellValue
{
    /// <summary>
    /// Used for null cell.
    /// Type is Undefined.
    /// </summary>
    public ExcelCellValue()
    {
        CellType = ExcelCellType.Undefined;
    }

    public ExcelCellValue(string value)
    {
        CellType = ExcelCellType.String;
        StringValue = value;
    }

    public ExcelCellValue(int value)
    {
        CellType = ExcelCellType.Integer;
        IntegerValue = value;
    }

    public ExcelCellValue(double value)
    {
        CellType = ExcelCellType.Double;
        DoubleValue = value;
    }

    public ExcelCellValue(DateOnly value)
    {
        CellType = ExcelCellType.DateOnly;
        DateOnlyValue = value;
    }

    public ExcelCellValue(DateTime value)
    {
        CellType = ExcelCellType.DateTime;
        DateTimeValue = value;
    }

    public ExcelCellValue(TimeOnly value)
    {
        CellType = ExcelCellType.TimeOnly;
        TimeOnlyValue = value;
    }

    /// <summary>
    /// type of the value of cell.
    /// The cell value can be empty.
    /// </summary>
    public ExcelCellType CellType { get; set; } = ExcelCellType.Undefined;

    /// <summary>
    /// Return true if the value of the cell is empty/blank.
    /// Type can be defined (string, integer,...) but in some cases, the type will be Undefined.
    /// </summary>
    public bool IsEmpty { get; set; } = false;

    /// <summary>
    /// part of the style/CellFormat.
    /// Used for DateTime, TimeSpan, DateOnly, TimeOnly formats.
    /// and also currency, percentage,... In this case, type is Double.
    /// </summary>
    public string? NumberFormat { get; set; } = null;

    /// <summary>
    /// part of the style/CellFormat.
    /// Number format id, if it exists. if not the default value is -1.
    /// </summary>
    public int NumberFormatId { get; set; } = -1;

    /// <summary>
    /// Set if the cell contains a formula.
    /// </summary>
    public string? Formula { get; set; } = null;

    /// <summary>
    /// Set if the type of the cell value is a string.
    /// </summary>
    public string? StringValue { get; set; } = null;

    /// <summary>
    /// Set if the type of the cell value is an integer.
    /// </summary>
    public int? IntegerValue { get; set; } = null;

    /// <summary>
    /// Set if the type of the cell value is a double.
    /// </summary>
    public double? DoubleValue { get; set; } = null;

    /// <summary>
    /// Set if the type of the cell value is a boolean.
    /// NOT-IMPLEMEMTED
    /// </summary>
    public bool? BooleanValue { get; set; } = null;

    /// <summary>
    /// Set if the type of the cell value is a DateOnly.
    /// </summary>
    public DateOnly? DateOnlyValue { get; set; } = null;

    /// <summary>
    /// Set if the type of the cell value is a DateTime.
    /// </summary>
    public DateTime? DateTimeValue { get; set; } = null;

    /// <summary>
    /// Set if the type of the cell value is a TimeOnly.
    /// </summary>
    public TimeOnly? TimeOnlyValue { get; set; } = null;

    //public TimeSpan? TimeSpanValue { get; set; }
}