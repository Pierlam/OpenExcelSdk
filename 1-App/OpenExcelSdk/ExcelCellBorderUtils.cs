using DocumentFormat.OpenXml.Spreadsheet;
using OpenExcelSdk.System;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcelSdk;

public class ExcelCellBorderUtils
{

    /// <summary>
    /// Create an excel cell border object from an OpenXML border object.
    /// </summary>
    /// <param name="openXmlBorder"></param>
    /// <returns></returns>
    public static ExcelCellBorder CreateCellBorder(object openXmlBorder)
    {
        ExcelCellBorderType borderType= ExcelCellBorderType.Left;
        ExcelCellBorderStyle borderStyle = ExcelCellBorderStyle.None;

        Color color= null;

        LeftBorder leftBorder= openXmlBorder as LeftBorder;
        if (leftBorder!=null)
        {
            borderType = ExcelCellBorderType.Left;
            if(leftBorder.Style!=null)
                borderStyle = ConvertBorderStyle(leftBorder.Style);

            //if (leftBorder.Color!= null)

            return new ExcelCellBorder(borderType, openXmlBorder, borderStyle);
        }

        RightBorder rightBorder= openXmlBorder as RightBorder;
        if (rightBorder!=null)
        {
            borderType = ExcelCellBorderType.Right;
            if (rightBorder.Style != null)
                borderStyle = ConvertBorderStyle(rightBorder.Style);

            return new ExcelCellBorder(borderType, openXmlBorder, borderStyle);
        }

        TopBorder topBorder= openXmlBorder as TopBorder;
        if (topBorder !=null)
        {
            borderType = ExcelCellBorderType.Top;
            if (topBorder.Style != null)
                borderStyle = ConvertBorderStyle(topBorder.Style);

            return new ExcelCellBorder(borderType, openXmlBorder, borderStyle);
        }

        BottomBorder bottomBorder= openXmlBorder as BottomBorder;   
        if (bottomBorder!=null)
        {
            borderType = ExcelCellBorderType.Bottom;
            if (bottomBorder.Style != null)
                borderStyle = ConvertBorderStyle(bottomBorder.Style);

            return new ExcelCellBorder(borderType, openXmlBorder, borderStyle);
        }

        DiagonalBorder diagonalBorder= openXmlBorder as DiagonalBorder;
        if (diagonalBorder != null)
        {
            borderType = ExcelCellBorderType.Diagonal;
            if (diagonalBorder.Style != null)
                borderStyle = ConvertBorderStyle(diagonalBorder.Style);

            return new ExcelCellBorder(borderType, openXmlBorder, borderStyle);
        }

        return null;
    }

    static ExcelCellBorderStyle ConvertBorderStyle(BorderStyleValues style)
    {
        if (style == null)
            return ExcelCellBorderStyle.None;

        if (style== BorderStyleValues.Thick)
            return ExcelCellBorderStyle.Thick;


        if (style == BorderStyleValues.Thin)
            return ExcelCellBorderStyle.Thin;

        if (style == BorderStyleValues.Hair )
            return ExcelCellBorderStyle.Hair;

        if (style == BorderStyleValues.DashDot)
            return ExcelCellBorderStyle.DashDot;

        if (style == BorderStyleValues.DashDotDot)
            return ExcelCellBorderStyle.DashDotDot;

        if (style == BorderStyleValues.Dotted)
            return ExcelCellBorderStyle.Dotted;

        if (style == BorderStyleValues.Dashed)
            return ExcelCellBorderStyle.Dashed;

        if (style == BorderStyleValues.Double)
            return ExcelCellBorderStyle.Double;

        if (style == BorderStyleValues.Medium)
            return ExcelCellBorderStyle.Medium;

        if (style == BorderStyleValues.MediumDashDot)
            return ExcelCellBorderStyle.MediumDashDotDot;

        if (style == BorderStyleValues.MediumDashDotDot)
            return ExcelCellBorderStyle.MediumDashDotDot;

        if (style == BorderStyleValues.MediumDashed)
            return ExcelCellBorderStyle.MediumDashDotDot;

        if (style == BorderStyleValues.SlantDashDot)
            return ExcelCellBorderStyle.SlantDashDot;

        return ExcelCellBorderStyle.None;
    }

}
