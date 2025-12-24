using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcelSdk;
public class Definitions
{
    /// <summary>
    /// Numbering format for Day-Month-Year with 4 digit year (e.g., 31/12/2024)
    /// </summary>
    public const string NumFmtDayMonthYear14 = "d/m/yyyy";

    /// <summary>
    /// NUmbering format for Euro currency with 2 decimal places (e.g., 1.234,56 €)
    /// </summary>
    public const string NumFmtEuro2Dec= "#,##0.00\\ \"€\"";
}
