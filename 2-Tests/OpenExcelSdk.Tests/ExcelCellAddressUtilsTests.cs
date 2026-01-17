using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcelSdk.Tests;

[TestClass]
public class ExcelCellAddressUtilsTests
{
    [TestMethod]
    public void CellAddressOk()
    {
        bool res= ExcelCellAddressUtils.GetColumnAndRowIndex("A2", out int colIdx, out int rowIdx);
        Assert.IsTrue(res);
        Assert.AreEqual(1, colIdx);
        Assert.AreEqual(2, rowIdx);

        res = ExcelCellAddressUtils.GetColumnAndRowIndex(" A2", out colIdx, out rowIdx);
        Assert.IsTrue(res);

        res = ExcelCellAddressUtils.GetColumnAndRowIndex(" A2 ", out colIdx, out rowIdx);
        Assert.IsTrue(res);

        res = ExcelCellAddressUtils.GetColumnAndRowIndex("AD28", out colIdx, out rowIdx);
        Assert.IsTrue(res);
        Assert.AreEqual(30, colIdx);
        Assert.AreEqual(28, rowIdx);


        res = ExcelCellAddressUtils.GetColumnAndRowIndex("ABC785", out colIdx, out rowIdx);
        Assert.IsTrue(res);
        Assert.AreEqual(731, colIdx);
        Assert.AreEqual(785, rowIdx);

        res = ExcelCellAddressUtils.GetColumnAndRowIndex("a5", out colIdx, out rowIdx);
        Assert.IsTrue(res);
        Assert.AreEqual(1, colIdx);
        Assert.AreEqual(5, rowIdx);

        res = ExcelCellAddressUtils.GetColumnAndRowIndex("aD64", out colIdx, out rowIdx);
        Assert.IsTrue(res);
        Assert.AreEqual(30, colIdx);
        Assert.AreEqual(64, rowIdx);

        // last col and row
        res = ExcelCellAddressUtils.GetColumnAndRowIndex("XFD1048576", out colIdx, out rowIdx);
        Assert.IsTrue(res);
        Assert.AreEqual(16384, colIdx);
        Assert.AreEqual(1048576, rowIdx);
    }

    [TestMethod]
    public void CellAddressErr()
    {
        bool res = ExcelCellAddressUtils.GetColumnAndRowIndex("A", out int colIdx, out int rowIdx);
        Assert.IsFalse(res);

        res = ExcelCellAddressUtils.GetColumnAndRowIndex("A+2", out colIdx, out rowIdx);
        Assert.IsFalse(res);

        res = ExcelCellAddressUtils.GetColumnAndRowIndex("A 2", out colIdx, out rowIdx);
        Assert.IsFalse(res);

        res = ExcelCellAddressUtils.GetColumnAndRowIndex("A2+", out colIdx, out rowIdx);
        Assert.IsFalse(res);

        res = ExcelCellAddressUtils.GetColumnAndRowIndex("A2-3", out colIdx, out rowIdx);
        Assert.IsFalse(res);
    }

    [TestMethod]
    public void TooHighErr()
    {
        bool res = ExcelCellAddressUtils.GetColumnAndRowIndex("XFE1", out int colIdx, out int rowIdx);
        Assert.IsFalse(res);


        res = ExcelCellAddressUtils.GetColumnAndRowIndex("A1048577", out colIdx, out rowIdx);
        Assert.IsFalse(res);


        res = ExcelCellAddressUtils.GetColumnAndRowIndex("XFE1048576", out colIdx, out rowIdx);
        Assert.IsFalse(res);
    }

}
