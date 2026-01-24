using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcelSdk.System.Export;

public class ExcelFontExport
{
    public ExcelFontExport(int sheetIndex, int fontId)
    {
        SheetIndex = sheetIndex;
        FontId = fontId;
    }

    public int SheetIndex { get; set; }
    public int FontId { get; set; }

}
