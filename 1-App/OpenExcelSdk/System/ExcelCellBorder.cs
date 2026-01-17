using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcelSdk.System;

public enum ExcelCellBorderType
{
    Left,
    Right,
    Top,
    Bottom,
    Diagonal
}

public enum ExcelCellBorderStyle
{
    None,
    Thin,
    Medium,
    Dashed,
    Dotted,
    Thick,
    Double,
    Hair,
    MediumDashed,
    DashDot,
    MediumDashDot,
    DashDotDot,
    MediumDashDotDot,
    SlantDashDot
}

/// <summary>
/// Excel Cell Border.
/// e.g. Left, Right, Top, Bottom, Diagonal border
/// </summary>
public class ExcelCellBorder
{
    public ExcelCellBorder(ExcelCellBorderType borderType , object openXmlBorder, ExcelCellBorderStyle borderStyle)
    {
        OpenXmlBorder= openXmlBorder;
        BorderType= borderType;
        BorderStyle= borderStyle;
    }

    public ExcelCellBorderType BorderType { get; set; }

    public object OpenXmlBorder { get; set; }

    public ExcelCellBorderStyle BorderStyle { get; set; }
}
