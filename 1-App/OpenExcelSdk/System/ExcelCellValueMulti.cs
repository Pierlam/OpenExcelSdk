using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcelSdk.System;


/// <summary>
/// Excel cell that can hold multiple values/types.
/// </summary>
public class ExcelCellValueMulti
{
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
        IntValue = value;
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

    public ExcelCellType CellType { get; set; }= ExcelCellType.Undefined;

    /// <summary>
    /// Used for DateTime, TimeSpan, DateOnly, TimeOnly formats. 
    /// and also currency, percentage,... In this case, type is Double.
    /// </summary>
    public string? DataFormat { get; set; } = null;

    public string? StringValue { get; set; } = null;

    public int? IntValue { get; set; } = null;

    public double? DoubleValue { get; set; } = null;
    
    public bool? BoolValue { get; set; } = null;

    public DateOnly? DateOnlyValue { get; set; }= null;

    public DateTime? DateTimeValue { get; set; } = null;

    public TimeOnly? TimeOnlyValue { get; set; } = null;

    //public TimeSpan? TimeSpanValue { get; set; }
}
