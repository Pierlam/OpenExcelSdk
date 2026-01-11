using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcelSdk.System.Export;

/// <summary>
/// All styles of an Excel file.
/// </summary>
public class ExcelStyles
{
    public List<SheetTable> ListSheetTable { get; set; } = new List<SheetTable>();  
    public List<StyleTable> ListStyleTable { get; set; } = new List<StyleTable>();
}
