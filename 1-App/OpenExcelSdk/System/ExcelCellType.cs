using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcelSdk.System;

/// <summary>
/// Type of a Excel cell value.
/// </summary>
public enum ExcelCellType
{
    Undefined,
    Error,
    Boolean,
    String,
    Integer,
    Double,
    DateOnly,
    DateTime,
    TimeOnly
}
