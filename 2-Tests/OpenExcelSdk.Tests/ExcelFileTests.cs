using OpenExcelSdk.Tests._50_Common;

namespace OpenExcelSdk.Tests;

/// <summary>
/// Open/Create Excel file tests.
/// </summary>
[TestClass]
public sealed class ExcelFileTests : TestBase
{
    [TestMethod]
    public void CreateExcelOk()
    {
        ExcelProcessor proc = new ExcelProcessor();

        string filename = PathFiles + "tmpCreation.xlsx";

        // remove the file from the previous test
        if (File.Exists(filename))
            File.Delete(filename);

        ExcelFile excelFile =proc.CreateExcelFile(filename);

        Assert.IsTrue(File.Exists(filename));

        proc.CloseExcelFile(excelFile);
    }

    [TestMethod]
    public void OpenExcelOk()
    {
        bool res;
        ExcelProcessor proc = new ExcelProcessor();

        string filename = PathFiles + "data3rows.xlsx";
        ExcelFile excelFile= proc.OpenExcelFile(filename);

        ExcelSheet excelSheet = proc.GetSheetAt(excelFile, 0);

        ExcelRow row = proc.GetRowAt(excelSheet, 0);
        Assert.IsNotNull(row);

        int lastRowIdx = proc.GetLastRowIndex(excelSheet);
        Assert.AreEqual(3, lastRowIdx);

        proc.CloseExcelFile(excelFile);
    }

    [TestMethod]
    public void OpenExcelNotExistsErr()
    {
        bool res;
        ExcelProcessor proc = new ExcelProcessor();
        try
        {
            string filename = PathFiles + "notexists.xlsx";
            ExcelFile excelFile = proc.OpenExcelFile(filename);
        }
        catch (ExcelException ex) 
        {
            Assert.AreEqual(ExcelErrorCode.FileNotFound, ex.ExcelErrorCode);
            return;
        }
        Assert.Fail("exception expected");
    }

    [TestMethod]
    public void OpenEmptyExcel()
    {
        bool res;
        ExcelProcessor proc = new ExcelProcessor();

        string filename = PathFiles + "empty.xlsx";
        ExcelFile excelFile = proc.OpenExcelFile(filename);

        ExcelSheet excelSheet = proc.GetSheetAt(excelFile, 0);

        ExcelRow row = proc.GetRowAt(excelSheet, 0);
        // no data row, not an error, just return null
        Assert.IsNull(row);

        int lastRowIdx = proc.GetLastRowIndex(excelSheet);
        Assert.AreEqual(0, lastRowIdx);


        // try to get a cell that does not exist -> should works
        ExcelCell cell = proc.GetCellAt(excelSheet, 2, 2);
        Assert.IsNull(cell);

        proc.CloseExcelFile(excelFile);
    }
}