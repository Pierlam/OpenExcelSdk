using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcelSdk.System;

public enum ExcelCellValueType
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
