using OpenExcelSdk.Tests._50_Common;

namespace OpenExcelSdk.Tests;

[TestClass]
public class SheetTests : TestBase
{
    [TestMethod]
    public void GetSheetByName()
    {
        ExcelProcessor proc = new ExcelProcessor();

        string filename = PathFiles + "hasManySheets.xlsx";
        ExcelFile excelFile = proc.OpenExcelFile(filename);

        ExcelSheet excelSheet = proc.GetSheetAt(excelFile, 0);
        Assert.IsNotNull(excelSheet);

        excelSheet = proc.GetSheetByName(excelFile, "Feuil1");
        Assert.IsNotNull(excelSheet);
        Assert.AreEqual("Feuil1", excelSheet.Sheet.Name.Value);

        excelSheet = proc.GetSheetByName(excelFile, "DoesNotExists");
        Assert.IsNull(excelSheet);

        proc.CloseExcelFile(excelFile);
    }

    [TestMethod]
    public void CreateSheet()
    {
        ExcelProcessor proc = new ExcelProcessor();

        string filename = PathFiles + "CreateSheet.xlsx";
        ExcelFile excelFile = proc.OpenExcelFile(filename);

        ExcelSheet excelSheet = proc.CreateSheet(excelFile, "mysheet");
        Assert.IsNotNull(excelSheet);

        try
        {
            // create sheet by the name is already used
            excelSheet = proc.CreateSheet(excelFile, "Sheet1");
        }
        catch (ExcelException ex)
        {
            Assert.AreEqual(ExcelErrorCode.UnableCreateSheet, ex.ExcelErrorCode);
            proc.CloseExcelFile(excelFile);
            return;
        }
        Assert.Fail("Exception expected");
    }
}