using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenExcelSdk.System;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcelSdk;
public class StyleMgr
{
    public bool GetCellNumberFormatId(ExcelSheet excelSheet, ExcelCell excelCell, out uint numFmtId)
    {
        numFmtId = 0;

        //--no style
        if (excelCell.Cell.StyleIndex == null) return true;

        var stylesPart = excelSheet.ExcelFile.WorkbookPart.WorkbookStylesPart;
        var cellFormat = (CellFormat)stylesPart.Stylesheet.CellFormats.ElementAt((int)excelCell.Cell.StyleIndex.Value);

        if (cellFormat.NumberFormatId == null) return true;

        numFmtId = cellFormat.NumberFormatId.Value;
        return true;
    }

    // Get number format string from formatId
    public bool GetNumberFormat(ExcelSheet excelSheet, uint formatId, out string dataFormat)
    {
        var stylesheet = excelSheet.ExcelFile.WorkbookPart.WorkbookStylesPart.Stylesheet;
        if (stylesheet.NumberingFormats != null)
        {
            foreach (NumberingFormat nf in stylesheet.NumberingFormats.Elements<NumberingFormat>())
            {
                if (nf.NumberFormatId.Value == formatId)
                {
                    dataFormat = nf.FormatCode.Value;
                    return true;
                }
            }
        }

        // Built-in format
        dataFormat=string.Empty;
        return false;
    }

}
