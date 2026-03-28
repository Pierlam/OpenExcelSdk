using OpenExcelSdk.Tests._50_Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Intrinsics.X86;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcelSdk.Tests;

/// <summary>
/// Get Excel Row tests.
/// </summary>
[TestClass]
public class GetRowTests: TestBase
{
    [TestMethod]
    public void GetRowOk()
    {
        bool res;
        ExcelProcessor proc = new ExcelProcessor();

        string filename = PathFiles + "get3rows.xlsx";
        ExcelFile excelFile = proc.OpenExcelFile(filename);

        ExcelSheet excelSheet = proc.GetSheetAt(excelFile, 0);

        // first row index is 1, so 0 is a wrong row index, but not an error, just return null
        ExcelRow row = proc.GetRowAtIndex(excelSheet, 0);
        Assert.IsNull(row);

        row = proc.GetRowAtIndex(excelSheet, 1);
        Assert.IsNotNull(row);

        row = proc.GetRowAtIndex(excelSheet, 2);
        Assert.IsNotNull(row);

        row = proc.GetRowAtIndex(excelSheet, 3);
        Assert.IsNotNull(row);

        // no more row, but not an error, just return null
        row = proc.GetRowAtIndex(excelSheet, 4);
        Assert.IsNull(row);

        int lastRowIdx = proc.GetLastRowIndex(excelSheet);
        Assert.AreEqual(3, lastRowIdx);

        proc.CloseExcelFile(excelFile);
    }

    [TestMethod]
    public void GetRowCellsCount()
    {
        bool res;
        ExcelProcessor proc = new ExcelProcessor();

        string filename = PathFiles + "getRowCellsCount.xlsx";
        ExcelFile excelFile = proc.OpenExcelFile(filename);

        ExcelSheet excelSheet = proc.GetSheetAt(excelFile, 0);

        // first row index is 1
        int count = proc.GetRowCellsCountAtIndex(excelSheet, 1);
        Assert.AreEqual(3, count);

        count = proc.GetRowCellsCountAtIndex(excelSheet, 2);
        Assert.AreEqual(5, count);

        count = proc.GetRowCellsCountAtIndex(excelSheet, 3);
        Assert.AreEqual(1, count);

        // row 4, no cell
        count = proc.GetRowCellsCountAtIndex(excelSheet, 4);
        Assert.AreEqual(0, count);

        proc.CloseExcelFile(excelFile);
    }

    [TestMethod]
    public void GetRowCellsRowIndex()
    {
        bool res;
        ExcelProcessor proc = new ExcelProcessor();

        string filename = PathFiles + "getRowCellsRowIndex.xlsx";
        ExcelFile excelFile = proc.OpenExcelFile(filename);

        ExcelSheet excelSheet = proc.GetSheetAt(excelFile, 0);

        // first row index is 1, 3 cells
        List<ExcelCell> listCells= proc.GetRowCellsAtAddress(excelSheet, 1);
        Assert.AreEqual(3, listCells.Count);

        // row #2, 5 cells
        listCells = proc.GetRowCellsAtAddress(excelSheet, 2);
        Assert.AreEqual(5, listCells.Count);

        // row #3, 1 cell
        listCells = proc.GetRowCellsAtAddress(excelSheet, 3);
        Assert.AreEqual(1, listCells.Count);

        // row #4, no cell
        listCells = proc.GetRowCellsAtAddress(excelSheet, 4);
        Assert.AreEqual(0, listCells.Count);

        // row -5, does not exist, but not an error, just return empty list
        listCells = proc.GetRowCellsAtAddress(excelSheet, -5);
        Assert.AreEqual(0, listCells.Count);

        proc.CloseExcelFile(excelFile);
    }

    [TestMethod]
    public void GetRowCellsExcelRow()
    {
        bool res;
        ExcelProcessor proc = new ExcelProcessor();

        string filename = PathFiles + "getRowCellsExcelRow.xlsx";
        ExcelFile excelFile = proc.OpenExcelFile(filename);

        ExcelSheet excelSheet = proc.GetSheetAt(excelFile, 0);

        // first row index is 1, 3 cells
        ExcelRow excelRow= proc.GetRowAtIndex(excelSheet, 1);
        Assert.IsNotNull(excelRow);
        List<ExcelCell> listCells = proc.GetRowCells(excelSheet, excelRow);
        Assert.AreEqual(3, listCells.Count);

        // row #2, 5 cells
        excelRow = proc.GetRowAtIndex(excelSheet, 2);
        Assert.IsNotNull(excelRow);
        listCells = proc.GetRowCells(excelSheet, excelRow);
        Assert.AreEqual(5, listCells.Count);

        // row #3, 1 cell
        excelRow = proc.GetRowAtIndex(excelSheet, 3);
        Assert.IsNotNull(excelRow);
        listCells = proc.GetRowCells(excelSheet, excelRow);
        Assert.AreEqual(1, listCells.Count);

        // row #4, no cell
        excelRow = proc.GetRowAtIndex(excelSheet, 4);
        Assert.IsNull(excelRow);

        // row -5, does not exist, but not an error, just return empty list
        excelRow = proc.GetRowAtIndex(excelSheet, -5);
        Assert.IsNull(excelRow);

        proc.CloseExcelFile(excelFile);
    }


    [TestMethod]
    public void GetRowAddressOk()
    {
        bool res;
        ExcelProcessor proc = new ExcelProcessor();

        string filename = PathFiles + "GetRowAddressOk.xlsx";
        ExcelFile excelFile = proc.OpenExcelFile(filename);

        ExcelSheet excelSheet = proc.GetSheetAt(excelFile, 0);

        // row #1: exists, row # 2 and #3, does not exist, row #4:exists, idx=2
        ExcelRow excelRow = proc.GetRowAtIndex(excelSheet, 2);

        int rowAddr= proc.GetRowAddress(excelSheet, excelRow);
        Assert.AreEqual(4, rowAddr);


        // row #3 does not exist, but not an error, just return 0
        excelRow = proc.GetRowAtIndex(excelSheet, 3);
        Assert.IsNull(excelRow);
    }

    [TestMethod]
    public void GetLastRowAddressOk()
    {
        bool res;
        ExcelProcessor proc = new ExcelProcessor();

        string filename = PathFiles + "GetLastRowAddressOk.xlsx";
        ExcelFile excelFile = proc.OpenExcelFile(filename);

        ExcelSheet excelSheet = proc.GetSheetAt(excelFile, 0);

        // row #1: exists, row # 2 and #3, does not exist, row #4:exists, idx=2
        int rowAddress = proc.GetLastRowAddress(excelSheet);
        Assert.AreEqual(4, rowAddress);

        int rowIdx = proc.GetLastRowIndex(excelSheet);
        Assert.AreEqual(2, rowIdx);
    }

}
