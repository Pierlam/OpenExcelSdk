using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcelSdk.Tests;

[TestClass]
public class ExcelUtilsTests
{
    [TestMethod]
    public void CellAddressOk()
    {
        bool res= ExcelUtils.GetColumnAndRowIndex("A2", out int colIdx, out int rowIdx);
        Assert.IsTrue(res);
        Assert.AreEqual(1, colIdx);
        Assert.AreEqual(2, rowIdx);

        res = ExcelUtils.GetColumnAndRowIndex(" A2", out colIdx, out rowIdx);
        Assert.IsTrue(res);

        res = ExcelUtils.GetColumnAndRowIndex(" A2 ", out colIdx, out rowIdx);
        Assert.IsTrue(res);

        res = ExcelUtils.GetColumnAndRowIndex("AD28", out colIdx, out rowIdx);
        Assert.IsTrue(res);
        Assert.AreEqual(30, colIdx);
        Assert.AreEqual(28, rowIdx);


        res = ExcelUtils.GetColumnAndRowIndex("ABC785", out colIdx, out rowIdx);
        Assert.IsTrue(res);
        Assert.AreEqual(731, colIdx);
        Assert.AreEqual(785, rowIdx);

        res = ExcelUtils.GetColumnAndRowIndex("a5", out colIdx, out rowIdx);
        Assert.IsTrue(res);
        Assert.AreEqual(1, colIdx);
        Assert.AreEqual(5, rowIdx);

        res = ExcelUtils.GetColumnAndRowIndex("aD64", out colIdx, out rowIdx);
        Assert.IsTrue(res);
        Assert.AreEqual(30, colIdx);
        Assert.AreEqual(64, rowIdx);

        // last col and row
        res = ExcelUtils.GetColumnAndRowIndex("XFD1048576", out colIdx, out rowIdx);
        Assert.IsTrue(res);
        Assert.AreEqual(16384, colIdx);
        Assert.AreEqual(1048576, rowIdx);
    }

    [TestMethod]
    public void CellAddressErr()
    {
        bool res = ExcelUtils.GetColumnAndRowIndex("A", out int colIdx, out int rowIdx);
        Assert.IsFalse(res);

        res = ExcelUtils.GetColumnAndRowIndex("A+2", out colIdx, out rowIdx);
        Assert.IsFalse(res);

        res = ExcelUtils.GetColumnAndRowIndex("A 2", out colIdx, out rowIdx);
        Assert.IsFalse(res);

        res = ExcelUtils.GetColumnAndRowIndex("A2+", out colIdx, out rowIdx);
        Assert.IsFalse(res);

        res = ExcelUtils.GetColumnAndRowIndex("A2-3", out colIdx, out rowIdx);
        Assert.IsFalse(res);
    }

    [TestMethod]
    public void TooHighErr()
    {
        bool res = ExcelUtils.GetColumnAndRowIndex("XFE1", out int colIdx, out int rowIdx);
        Assert.IsFalse(res);


        res = ExcelUtils.GetColumnAndRowIndex("A1048577", out colIdx, out rowIdx);
        Assert.IsFalse(res);


        res = ExcelUtils.GetColumnAndRowIndex("XFE1048576", out colIdx, out rowIdx);
        Assert.IsFalse(res);
    }

}
