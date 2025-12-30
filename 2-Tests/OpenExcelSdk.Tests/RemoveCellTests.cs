using OpenExcelSdk.Tests._50_Common;

namespace OpenExcelSdk.Tests;

[TestClass]
public class RemoveCellTests : TestBase
{
    [TestMethod]
    public void SetCellValueString()
    {
        bool res;
        ExcelProcessor proc = new ExcelProcessor();

        string filename = PathFiles + "RemoveCell.xlsx";
        ExcelFile excelFile = proc.OpenExcelFile(filename);

        ExcelSheet excelSheet = proc.GetSheetAt(excelFile, 0);

        ExcelCell cell;
        ExcelCellValueMulti cellValueMulti;

        //--B2: already null!
        res = proc.RemoveCell(excelSheet, "B2");
        Assert.IsTrue(res);

        //--B3:
        res = proc.RemoveCell(excelSheet, "B3");
        Assert.IsTrue(res);

        //--B4:
        res = proc.RemoveCell(excelSheet, 2, 4);
        Assert.IsTrue(res);

        //--B5:
        res = proc.RemoveCell(excelSheet, 2, 5);
        Assert.IsTrue(res);

        // save the changes
        proc.CloseExcelFile(excelFile);

        //==>check the excel content
        excelFile = proc.OpenExcelFile(filename);
        excelSheet = proc.GetSheetAt(excelFile, 0);

        //--B2: null
        cell = proc.GetCellAt(excelSheet, 2, 2);
        Assert.IsNull(cell);

        //--B3:
        cell = proc.GetCellAt(excelSheet, 2, 3);
        Assert.IsNull(cell);
    }
}