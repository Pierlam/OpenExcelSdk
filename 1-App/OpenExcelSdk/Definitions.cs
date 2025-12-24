using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcelSdk;

/// <summary>
/// Some general definitions.
/// 
/// ---
/// cell numbering format Help:
/// https://stackoverflow.com/questions/36670768/openxml-cell-datetype-is-null
/// https://stackoverflow.com/questions/4655565/reading-dates-from-openxml-excel-files
/// 
/// ECMA-376, Second Edition, Part 1 - Fundamentals And Markup Language Reference section 18.8.30 page 1786:
/// https://www.ecma-international.org/publications-and-standards/standards/ecma-376/
///
/// </summary>
public class Definitions
{
    /// <summary>
    /// Numbering format for Integer (e.g., 1234)
    /// Built-in format id: 1
    /// </summary>
    public const string NumFmtInteger2 = "0";

    /// <summary>
    /// Numbering format for Number with 2 decimal places (e.g., 1234.56)
    /// Built-in format id: 2
    /// </summary>
    public const string NumFmtNumberTwoDec2 = "0.00";

    /// <summary>
    /// Numbering format percent for integer.
    /// Built-in format id: 9
    /// </summary>
    public const string NumFmtNumberPercentInt9 = "0%";

    /// <summary>
    /// Numbering format percent with 2 decimal places.
    /// Built-in format id: 10
    /// </summary>
    public const string NumFmtNumberPercent10 = "0.00%";

    /// <summary>
    /// Numbering format for Day-Month-Year with 4 digit year (e.g., 31/12/2024)
    /// Built-in format id: 14
    /// </summary>
    public const string NumFmtDayMonthYear14 = "d/m/yyyy";

    /// <summary>
    /// Numbering format for Hour-Minute (e.g., 14:30)
    /// Built-in format id: 20
    /// </summary>
    public const string NumFmtHourMin20 = "h:mm";

    /// <summary>
    /// Numbering format for Hour-Minute-Second (e.g., 14:30:45)
    /// Built-in format id: 21
    /// </summary>
    public const string NumFmtHourMinSec21 = "h:mm:ss";

    /// <summary>
    /// Numbering format for currency. General case.
    /// Built-in format id: 44
    /// </summary>
    public const string NumFmtCurrencyGeneral44 = "_(\"$\"* #,##0.00_);_(\"$\"* \\(#,##0.00\\);_(\"$\"* \"-\"??_);_(@_)";

    /// <summary>
    /// Numbering format for DateTime with long format (e.g., 10/12/2025 12:34:56)
    /// Custom format.
    /// </summary>
    public const string NumFmtDateTimeLong = "dd/mm/yyyy\\ hh:mm:ss";

    /// <summary>
    /// Numbering format for Euro currency with 2 decimal places (e.g., 1.234,56 €)
    /// Custom format.
    /// </summary>
    public const string NumFmtEuroTwoDec= "#,##0.00\\ \"€\"";
}
