using OpenExcelSdk.System.Export;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcelSdk.Export;

public class ExcelAbstractExporter
{
    public static void Export(ExcelProcessor excelProcessor, ExcelAllStylesExport excelStyles, ExcelFile excelFileIn, ExcelFile excelFileOut)
    {
        // the first sheet exists already
        ExcelSheet excelSheet = excelProcessor.GetSheetAt(excelFileOut, 0);

        int numline = 1;
        excelProcessor.SetCellValue(excelSheet, "A" + numline.ToString(), "Title");
        excelProcessor.SetCellValue(excelSheet, "B" + numline.ToString(), "OpenExcelSdk - Export styles");

        numline++;
        excelProcessor.SetCellValue(excelSheet, "A" + numline.ToString(), "Version");
        excelProcessor.SetCellValue(excelSheet, "B" + numline.ToString(), 1);

        numline++;
        excelProcessor.SetCellValue(excelSheet, "A" + numline.ToString(), "Ref");
        excelProcessor.SetCellValue(excelSheet, "B" + numline.ToString(), "https://www.nuget.org/packages/OpenExcelSdk");

        numline++;
        excelProcessor.SetCellValue(excelSheet, "A" + numline.ToString(), "Fullfilename");
        excelProcessor.SetCellValue(excelSheet, "B" + numline.ToString(), Path.GetFullPath(excelFileIn.Filename));

        numline++;
        excelProcessor.SetCellValue(excelSheet, "A" + numline.ToString(), "Filename");
        excelProcessor.SetCellValue(excelSheet, "B" + numline.ToString(), Path.GetFileName(excelFileIn.Filename));

        numline++;
        excelProcessor.SetCellValue(excelSheet, "A" + numline.ToString(), "Sheet count");
        excelProcessor.SetCellValue(excelSheet, "B" +numline.ToString(), excelProcessor.GetSheetCount(excelFileIn));

        // get cells count
        numline++;
        excelProcessor.SetCellValue(excelSheet, "A" + numline.ToString(), "Cells count");
        excelProcessor.SetCellValue(excelSheet, "B" + numline.ToString(), excelStyles.CellsTotalCount);

        // get cells count
        numline++;
        int nbcell = excelStyles.CellsTotalCount;
        if(nbcell>excelStyles.CellsMaxLoadCount)nbcell= excelStyles.CellsMaxLoadCount;
        excelProcessor.SetCellValue(excelSheet, "A" + numline.ToString(), "Cells loaded");
        excelProcessor.SetCellValue(excelSheet, "B" + numline.ToString(), nbcell);

        // get shared strings count
        numline++;
        excelProcessor.SetCellValue(excelSheet, "A" + numline.ToString(), "SharedStrings count");
        excelProcessor.SetCellValue(excelSheet, "B" + numline.ToString(), excelStyles.SharedStringsTotalCount);

        numline++;
        excelProcessor.SetCellValue(excelSheet, "A" + numline.ToString(), "SharedStrings loaded");
        excelProcessor.SetCellValue(excelSheet, "B" + numline.ToString(), excelStyles.ListSharedStrings.Count);

        numline++;
        excelProcessor.SetCellValue(excelSheet, "A" + numline.ToString(), "Styles count");
        excelProcessor.SetCellValue(excelSheet, "B" + numline.ToString(), excelStyles.ListStyles.Count);

        numline++;
        excelProcessor.SetCellValue(excelSheet, "A" + numline.ToString(), "Fills count");
        excelProcessor.SetCellValue(excelSheet, "B" + numline.ToString(), excelStyles.ListFills.Count);

        numline++;
        excelProcessor.SetCellValue(excelSheet, "A" + numline.ToString(), "Borders count");
        excelProcessor.SetCellValue(excelSheet, "B" + numline.ToString(), excelStyles.ListBorders.Count);

        numline++;
        excelProcessor.SetCellValue(excelSheet, "A" + numline.ToString(), "Fonts count");
        excelProcessor.SetCellValue(excelSheet, "B" + numline.ToString(), excelStyles.ListFonts.Count);

        numline++;
        excelProcessor.SetCellValue(excelSheet, "A" + numline.ToString(), "Errors count");
        excelProcessor.SetCellValue(excelSheet, "B" + numline.ToString(), excelStyles.ListError.Count);

        for(int i=0;i< excelStyles.ListError.Count;i++)
        {
            numline++;
            excelProcessor.SetCellValue(excelSheet, "A" + numline.ToString(), "Error #" + (i + 1).ToString());
            excelProcessor.SetCellValue(excelSheet, "B" + numline.ToString(), excelStyles.ListError[i]);
        }
    }
}
