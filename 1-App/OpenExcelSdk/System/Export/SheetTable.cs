using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcelSdk.System.Export;

public class SheetTable
{
    public SheetTable(int index, string name)
    {
        Index = index;
        Name = name;
    }

    public int Index { get; set; }
    public string Name { get; set; }

}
