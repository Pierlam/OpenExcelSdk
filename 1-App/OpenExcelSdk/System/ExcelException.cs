using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcelSdk;

/// <summary>
/// Specific excel exception.
/// </summary>
public class ExcelException :Exception
{
    public ExcelException()
    {
    }

    public ExcelException(string message)
        : base(message)
    {
    }

    public ExcelException(string message, Exception inner)
        : base(message, inner)
    {
    }
}
