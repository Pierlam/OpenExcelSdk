using OpenExcelSdk.System;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcelSdk;
public class BuiltInNumberFormatMgr
{
    /// <summary>
    /// Get the built-in format id for the given format string.
    /// </summary>
    /// <param name="format"></param>
    /// <param name="formatId"></param>
    /// <returns></returns>
    public static bool GetFormatId(string format, out uint formatId)
    {
        formatId = 0;
        if (string.IsNullOrEmpty(format))
            return false;

        if(format=="0")
        {
            formatId = 1;
            return true;
        }
        if (format == "0.00")
        {
            formatId = 2;
            return true;
        }
        if (format == "#,##0")
        {
            formatId = 3;
            return true;
        }
        if (format == "#,##0.00")
        {
            formatId = 4;
            return true;
        }
        if (format == "0%")
        {
            formatId = 9;
            return true;
        }
        if (format == "0.00 %")
        {
            formatId = 10;
            return true;
        }
        if (format == "0.00E+00")
        {
            formatId = 11;
            return true;
        }
        if (format == "# ?/?")
        {
            formatId = 12;
            return true;
        }

        if (format == "# ??/??")
        {
            formatId = 13;
            return true;
        }
        if (format == "d/m/yyyy")
        {
            formatId = 14;
            return true;
        }
        if (format == "d-mmm-yy")
        {
            formatId = 15;
            return true;
        }
        if (format == "d-mmm")
        {
            formatId = 16;
            return true;
        }
        if (format == "mmm-yy")
        {
            formatId = 17;
            return true;
        }
        if (format == "h:mm AM/PM")
        {
            formatId = 18;
            return true;
        }
        if (format == "h:mm:ss AM/PM")
        {
            formatId = 19;
            return true;
        }
        if (format == "h:mm")
        {
            formatId = 20;
            return true;
        }
        if (format == "h:mm:ss")
        {
            formatId = 21;
            return true;
        }
        if (format == "m/d/yyyy h:mm")
        {
            formatId = 22;
            return true;
        }

        // TODO:
        // 27 = '[$-404]e/m/d'
        // 28 = [$-404]e"?"m"?"d"?" m"?"d"?"
        // 30 = 'm/d/yy'
        // 36 = '[$-404]e/m/d'

        //---
        // 37 = '#,##0 ;(#,##0)'               ou "#,##0_);(#,##0)"
        // 38 = '#,##0 ;[Red](#,##0)'          ou "#,##0_);[Red]"
        // 39 = '#,##0.00;(#,##0.00)'          ou "#,##0.00_);(#,##0.00)"
        // 40 = '#,##0.00;[Red](#,##0.00)'     ou  "#,##0.00_);[Red]"


        if (format == "_(\"$\"* #,##0.00_);_(\"$\"* \\(#,##0.00\\);_(\"$\"* \"-\"??_);_(@_)")
        {
            formatId = 44;
            return true;
        }

        // 45 = 'mm:ss'
        // 46 = '[h]:mm:ss'
        // 47 = 'mmss.0'       ou "mm:ss.0"
        // 48 = '##0.0E+0'
        // 49 = '@'


        // 50 = '[$-404]e/m/d'
        // 55 = 'yyyy/mm/dd'
        // 57 = '[$-404]e/m/d'
        // 59 = 't0'
        // 60 = 't0.00'
        // 61 = 't#,##0'
        // 62 = 't#,##0.00'
        // 67 = 't0%'
        // 68 = 't0.00%'
        // 69 = 't# ?/?'
        // 70 = 't# ??/??'          

        return false;
    }

    /// <summary>
    /// Returns the built-in number format string for the given numFmtId.
    /// https://github.com/ClosedXML/ClosedXML/wiki/NumberFormatId-Lookup-Table
    /// Excel built-in date formats are typically between 14 and 22 (and also 45, 46, 47 ??)
    /// </summary>
    /// <param name="numFmtId"></param>
    /// <returns></returns>
    public static bool GetFormatAndType(uint numFmtId, out string numberFormat, out ExcelCellType cellType)
    {
        numberFormat = string.Empty;
        cellType= ExcelCellType.Undefined;

        // General, can be string or number
        if (numFmtId == 0)
            return false;

        //  1 = '0'
        if (numFmtId == 1)
        {
            numberFormat = "0";
            cellType = ExcelCellType.Double;
            return true;
        }

        if (numFmtId == 2)
        {
            numberFormat = "0.00";
            cellType = ExcelCellType.Double;
            return true;
        }

        //--3: #,##0
        if (numFmtId == 3)
        {
            numberFormat = "#,##0";
            cellType = ExcelCellType.Integer;
            return true;
        }

        //--4: #,##0.00
        if (numFmtId == 4)
        {
            numberFormat = "#,##0.00";
            cellType = ExcelCellType.Double;
            return true;
        }

        /* TODO: currency formats
         5 = '$#,##0;\-$#,##0'
         6 = '$#,##0;[Red]\-$#,##0'
         7 = '$#,##0.00;\-$#,##0.00'
         8 = '$#,##0.00;[Red]\-$#,##0.00'
         */

        //--9: 0 %
        if (numFmtId == 9)
        {
            // displpayed an integer but the value can be a double
            numberFormat = "0%";
            cellType = ExcelCellType.Double;
            return true;
        }

        //--10: 0.00 %
        if (numFmtId == 10)
        {
            numberFormat = "0.00 %";
            cellType = ExcelCellType.Double;
            return true;
        }

        //--11: 0.00E+00
        if (numFmtId == 11)
        {
            numberFormat = "0.00E+00";
            cellType = ExcelCellType.Double;
            return true;
        }

        //--12: # ?/?
        if (numFmtId == 12)
        {
            numberFormat = "# ?/?";
            cellType = ExcelCellType.Double;
            return true;
        }

        //--13: # ??/??
        if (numFmtId == 13)
        {
            numberFormat = "# ??/??";
            cellType = ExcelCellType.Double;
            return true;
        }

        if (numFmtId == 14)
        {
            numberFormat = "d/m/yyyy";
            cellType = ExcelCellType.DateOnly;
            return true;
        }

        // 15 = 'd-mmm-yy'
        if (numFmtId == 15)
        {
            numberFormat = "d-mmm-yy";
            cellType = ExcelCellType.DateOnly;
            return true;
        }

        if (numFmtId == 16)
        {
            numberFormat = "d-mmm";
            cellType = ExcelCellType.DateOnly;
            return true;
        }

        //17 = 'mmm-yy'
        if (numFmtId == 17)
        {
            numberFormat = "mmm-yy";
            cellType = ExcelCellType.DateOnly;
            return true;
        }

        //18 = 'h:mm AM/PM'
        if (numFmtId == 18)
        {
            numberFormat = "h:mm AM/PM";
            cellType = ExcelCellType.TimeOnly;
            return true;
        }

        //19 = 'h:mm:ss AM/PM'
        if (numFmtId == 19)
        {
            numberFormat = "h:mm:ss AM/PM";
            cellType = ExcelCellType.TimeOnly;
            return true;
        }

        //20 = 'h:mm'
        if (numFmtId == 20)
        {
            numberFormat = "h:mm";
            cellType = ExcelCellType.TimeOnly;
            return true;
        }

        //21 = 'h:mm:ss'
        if (numFmtId == 21)
        {
            numberFormat = "h:mm:ss";
            cellType = ExcelCellType.TimeOnly;
            return true;
        }

        // 22 = "m/d/yyyy h:mm"
        if (numFmtId == 22)
        {
            numberFormat = "m/d/yyyy h:mm";
            cellType = ExcelCellType.DateTime;
            return true;
        }

        // TODO:
        // 27 = '[$-404]e/m/d'
        // 28 = [$-404]e"?"m"?"d"?" m"?"d"?"
        // 30 = 'm/d/yy'
        // 36 = '[$-404]e/m/d'

        //---
        // 37 = '#,##0 ;(#,##0)'               ou "#,##0_);(#,##0)"
        // 38 = '#,##0 ;[Red](#,##0)'          ou "#,##0_);[Red]"
        // 39 = '#,##0.00;(#,##0.00)'          ou "#,##0.00_);(#,##0.00)"
        // 40 = '#,##0.00;[Red](#,##0.00)'     ou  "#,##0.00_);[Red]"


        // 44 = '_("$"* #,##0.00_);_("$"* \(#,##0.00\);_("$"* "-"??_);_(@_)'  -> currency with 2 decimals
        if (numFmtId == 44)
        {
            numberFormat = "_(\"$\"* #,##0.00_);_(\"$\"* \\(#,##0.00\\);_(\"$\"* \"-\"??_);_(@_)";
            cellType = ExcelCellType.Double;
            return true;
        }

        // 45 = 'mm:ss'
        // 46 = '[h]:mm:ss'
        // 47 = 'mmss.0'       ou "mm:ss.0"
        // 48 = '##0.0E+0'
        // 49 = '@'


        // 50 = '[$-404]e/m/d'
        // 55 = 'yyyy/mm/dd'
        // 57 = '[$-404]e/m/d'
        // 59 = 't0'
        // 60 = 't0.00'
        // 61 = 't#,##0'
        // 62 = 't#,##0.00'
        // 67 = 't0%'
        // 68 = 't0.00%'
        // 69 = 't# ?/?'
        // 70 = 't# ??/??'          


        // not a built-in data format
        numberFormat = string.Empty;
        cellType = ExcelCellType.Undefined;
        return false;
    }
}
