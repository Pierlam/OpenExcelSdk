using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcelSdk.System;
public class ExcelRow
{
    public ExcelRow(Row row)
    {
        Row = row;
    }

    /// <summary>
    /// Open Xml row object.
    /// </summary>
    public Row Row { get; set; }
}
