using OpenExcelSdk.Tests._50_Common;

namespace OpenExcelSdk.Tests;

[TestClass]
public class SheetTests : TestBase
{
    [TestMethod]
    public void GetSheetByName()
    {
        bool res;
        ExcelError error;
        ExcelProcessor proc = new ExcelProcessor();

        string filename = PathFiles + "hasManySheets.xlsx";
        ExcelFile excelFile = proc.OpenExcelFile(filename);

        res = proc.GetSheetAt(excelFile, 0, out ExcelSheet excelSheet, out error);
        Assert.IsTrue(res);
        Assert.IsNull(error);

        res = proc.GetSheetByName(excelFile, "Feuil1", out excelSheet, out error);
        Assert.IsTrue(res);
        Assert.IsNull(error);
        Assert.AreEqual("Feuil1", excelSheet.Sheet.Name.Value);

        res = proc.GetSheetByName(excelFile, "DoesNotExists", out excelSheet, out error);
        Assert.IsFalse(res);
        Assert.IsNull(error);

        proc.CloseExcelFile(excelFile);
    }

    [TestMethod]
    public void CreateSheet()
    {
        bool res;
        ExcelError error;
        ExcelProcessor proc = new ExcelProcessor();

        string filename = PathFiles + "CreateSheet.xlsx";
        ExcelFile excelFile = proc.OpenExcelFile(filename);

        res = proc.CreateSheet(excelFile, "mysheet", out ExcelSheet excelSheet, out error);
        Assert.IsTrue(res);

        // create sheet by the name is already used
        res = proc.CreateSheet(excelFile, "Sheet1", out excelSheet, out error);
        Assert.IsFalse(res);

        proc.CloseExcelFile(excelFile);
    }
}