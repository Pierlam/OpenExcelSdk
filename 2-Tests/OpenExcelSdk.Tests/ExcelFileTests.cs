using OpenExcelSdk.System;
using OpenExcelSdk.Tests._50_Common;

namespace OpenExcelSdk.Tests;

/// <summary>
/// Open/Create Excel file tests.
/// </summary>
[TestClass]
public sealed class ExcelFileTests :TestBase
{
    [TestMethod]
    public void CreateExcelOk()
    {
        bool res;
        ExcelError error;
        ExcelProcessor proc = new ExcelProcessor();

        string filename = PathFiles + "tmpCreation.xlsx";

        // remove the file from the previous test
        if (File.Exists(filename))
            File.Delete(filename);

        res=proc.CreateExcelFile(filename, out ExcelFile excelFile, out error);
        Assert.IsTrue(res);
        Assert.IsNull(error);

        Assert.IsTrue(File.Exists(filename));

        res = proc.Close(excelFile, out error);
        Assert.IsTrue(res);
        Assert.IsNull(error);
    }

    [TestMethod]
    public void OpenExcelOk()
    {
        bool res;
        ExcelError error;
        ExcelProcessor proc = new ExcelProcessor();

        string filename = PathFiles+ "data3rows.xlsx";
        res=proc.Open(filename, out ExcelFile excelFile, out error);
        Assert.IsTrue(res);
        Assert.IsNull(error);

        res=proc.GetSheetAt(excelFile, 0, out ExcelSheet excelSheet, out error);
        Assert.IsTrue(res);
        Assert.IsNull(error);

        res = proc.GetRowAt(excelSheet, 0, out ExcelRow row, out error);
        Assert.IsTrue(res);
        Assert.IsNull(error);

        int lastRowIdx = proc.GetLastRowIndex(excelSheet);
        Assert.AreEqual(3, lastRowIdx);

        res= proc.Close(excelFile, out error);
        Assert.IsTrue(res);
        Assert.IsNull(error);
    }
}
