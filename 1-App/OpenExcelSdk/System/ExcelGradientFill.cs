using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenExcelSdk.System;

public class ExcelGradientFill
{
    public ExcelGradientFill(GradientFill gradientFill) 
    {
        GradientFill = gradientFill;
    }


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
