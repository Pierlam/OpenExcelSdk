using DocumentFormat.OpenXml.Packaging;
using OpenExcelSdk.System.Export;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcelSdk.Export;

public class ExcelSharedStringExporter
{
    public static void Export(ExcelProcessor excelProcessor, ExcelAllStylesExport excelStyles, ExcelFile excelFileOut)
    {
        ExcelSheet excelSheetOut = excelProcessor.CreateSheet(excelFileOut, "SharedStrings");

        // create the out header
        CreateOutHeader(excelProcessor, excelSheetOut);

        for (int i =0; i< excelStyles.ListSharedStrings.Count;i++)
        {
            string rowIdx = (i + 2).ToString();

            excelProcessor.SetCellValue(excelSheetOut, "A" + rowIdx, excelStyles.ListSharedStrings[i].Index);
            excelProcessor.SetCellValue(excelSheetOut, "B" + rowIdx, excelStyles.ListSharedStrings[i].Text);
        }
    }


    /// <summary>
    /// Create the out header
    /// </summary>
    /// <param name="proc"></param>
    /// <param name="excelSheet"></param>
    static void CreateOutHeader(ExcelProcessor proc, ExcelSheet excelSheet)
    {
        proc.SetCellValue(excelSheet, "A1", "Idx");
        proc.SetCellValue(excelSheet, "B1", "Text");

    }

}
