using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcelSdk.System;
public class ExcelCell
{
    public ExcelCell(ExcelSheet excelSheet, Cell cell)
    {
        ExcelSheet= excelSheet;
        Cell = cell;
    }

    /// <summary>
    /// The sheet where is placed the cell.
    /// </summary>
    public ExcelSheet ExcelSheet { get; set; }

    /// <summary>
    /// Open Xml cell object.
    /// </summary>
    public  Cell Cell { get; set; }
}
