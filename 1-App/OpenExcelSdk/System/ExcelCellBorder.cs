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

/// <summary>
/// Excel Cell Border.
/// e.g. Left, Right, Top, Bottom, Diagonal border
/// </summary>
public class ExcelCellBorder
{
    public ExcelCellBorder(object openXmlBorder)
    {
        OpenXmlBorder= openXmlBorder;
        if (openXmlBorder is LeftBorder)
        { 
            BorderType = ExcelCellBorderType.Left;
             //((LeftBorder)openXmlBorder).Style
        }

        if (openXmlBorder is RightBorder) BorderType = ExcelCellBorderType.Right;
        if (openXmlBorder is TopBorder) BorderType = ExcelCellBorderType.Top;
        if (openXmlBorder is BottomBorder) BorderType = ExcelCellBorderType.Bottom;
        if (openXmlBorder is DiagonalBorder) BorderType = ExcelCellBorderType.Diagonal;
    }

    public ExcelCellBorderType BorderType { get; set; }

    public object OpenXmlBorder { get; set; }
}
