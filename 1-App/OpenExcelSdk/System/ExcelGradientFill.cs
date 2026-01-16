using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcelSdk.System;

public class ExcelGradientFill
{
    public ExcelGradientFill() { }


    /// <summary>
    /// OpenXml object.
    /// </summary>
    public GradientFill GradientFill { get; set; }

    public double Degree     
    {
        get
        {
            if (GradientFill == null) return 0;
            return GradientFill.Degree != null ? GradientFill.Degree.Value : 0;
        }
    }
}
