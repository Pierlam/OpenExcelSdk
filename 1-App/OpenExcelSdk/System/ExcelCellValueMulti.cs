namespace OpenExcelSdk;

/// <summary>
/// Excel cell that can hold multiple values/types.
/// </summary>
public class ExcelCellValueMulti
{
    /// <summary>
    /// Used for null cell.
    /// </summary>
    public ExcelCellValueMulti()
    {
        CellType = ExcelCellType.Undefined;
    }

    public ExcelCellValueMulti(string value)
    {
        CellType = ExcelCellType.String;
        StringValue = value;
    }

    public ExcelCellValueMulti(int value)
    {
        CellType = ExcelCellType.Integer;
        IntegerValue = value;
    }

    public ExcelCellValueMulti(double value)
    {
        CellType = ExcelCellType.Double;
        DoubleValue = value;
    }

    public ExcelCellValueMulti(DateOnly value)
    {
        CellType = ExcelCellType.DateOnly;
        DateOnlyValue = value;
    }

    public ExcelCellValueMulti(DateTime value)
    {
        CellType = ExcelCellType.DateTime;
        DateTimeValue = value;
    }

    public ExcelCellValueMulti(TimeOnly value)
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
    /// Type can be defined, or not, it will be Undefined in this case.
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