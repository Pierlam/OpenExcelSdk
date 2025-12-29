using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcelSdk;

/// <summary>
/// To manage shared string.
/// </summary>
public class SharedStringMgr
{

    public static bool GetSharedStringValue(ExcelSheet excelSheet, ExcelCell excelCell, out string stringValue)
    {
        stringValue = string.Empty;

        if (excelCell.Cell.DataType.Value != CellValues.SharedString) return false;

        // For shared strings, look up the value in the shared strings table.
        var stringTable = excelSheet.ExcelFile.WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

        // If the shared string table is missing, something is wrong.
        // Return the index that is in the cell.
        // Otherwise, look up the correct text in the table.
        if (stringTable is not null)
        {
            string value = excelCell.Cell.CellValue.InnerText;
            stringValue = stringTable.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
            return true;
        }
        return false;
    }

    /// <summary>
    /// Get the SharedStringTablePart. If it does not exist, create a new one.
    /// </summary>
    /// <param name="excelFile"></param>
    /// <returns></returns>
    public static SharedStringTablePart GetOrCreateSharedStringTablePart(WorkbookPart workbookPart)
    {
        if (workbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
            return workbookPart.GetPartsOfType<SharedStringTablePart>().First();

        return workbookPart.AddNewPart<SharedStringTablePart>();
    }

    /// <summary>
    /// Given text and a SharedStringTablePart, creates a SharedStringItem with the specified text
    ///  and inserts it into the SharedStringTablePart. If the item already exists, returns its index.
    /// </summary>
    /// <param name="text"></param>
    /// <param name="shareStringPart"></param>
    /// <returns></returns>
    public static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
    {
        // If the part does not contain a SharedStringTable, create one.
        shareStringPart.SharedStringTable ??= new SharedStringTable();

        int i = 0;

        // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
        foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
        {
            if (item.InnerText == text)
                return i;

            i++;
        }

        // The text does not exist in the part. Create the SharedStringItem and return its index.
        shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new Text(text)));

        return i;
    }

}
