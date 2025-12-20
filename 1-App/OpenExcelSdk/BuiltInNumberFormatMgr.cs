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
    /// Returns the built-in number format string for the given numFmtId.
    /// https://github.com/ClosedXML/ClosedXML/wiki/NumberFormatId-Lookup-Table
    /// Excel built-in date formats are typically between 14 and 22 (and also 45, 46, 47 ??)
    /// </summary>
    /// <param name="numFmtId"></param>
    /// <returns></returns>
    public static bool Get(uint numFmtId, out string numberFormat, out ExcelCellType cellType)
    {
        numberFormat = string.Empty;
        cellType= ExcelCellType.Undefined;

        // General, can be string or number
        if (numFmtId == 0)
            return false;

        if (numFmtId == 1)
        {

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

        // not a built-in data format
        numberFormat = string.Empty;
        cellType = ExcelCellType.Undefined;
        return false;
    }
}
