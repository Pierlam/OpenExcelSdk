using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcelSdk.System.Export;

public class ExcelSharedStringExport
{
    public ExcelSharedStringExport(int index, string text)
    {
        Index = index;
        Text = text;
    }
    public int Index { get; set; }
    public string Text { get; set; }    
}
