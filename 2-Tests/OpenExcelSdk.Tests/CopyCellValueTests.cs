using DocumentFormat.OpenXml.Spreadsheet;
using OpenExcelSdk.System;
using OpenExcelSdk.Tests._50_Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcelSdk.Tests;

[TestClass]
public class CopyCellValueTests : TestBase
{
    /// <summary>
    /// From a excel/sheet source to another excel/sheet destination, copy the value of a cell to another cell.
    /// </summary>
    [TestMethod]
    public void CopyCellValueStd()
    {
        bool res;
        ExcelProcessor proc = new ExcelProcessor();

        string filename = PathFiles + "CopyCellValue.xlsx";
        ExcelFile excelFile = proc.OpenExcelFile(filename);
        ExcelSheet excelSheet = proc.GetSheetAt(excelFile, 0);

        string filenameDest = PathFiles + "CopyCellValueDest.xlsx";
        ExcelFile excelFileDest = proc.OpenExcelFile(filenameDest);
        ExcelSheet excelSheetDest = proc.GetSheetAt(excelFileDest, 0);

        ExcelCell cell;
        ExcelCellValue excelCellValue;

        //--A2 -> B2: text, target cell is empty
        res = proc.CopyCellValue(excelSheet, "A2", excelSheetDest, "B2");
        Assert.IsTrue(res);

        //--A3 -> B3: 12, target cell is empty
        res = proc.CopyCellValue(excelSheet, "A3", excelSheetDest, "B3");
        Assert.IsTrue(res);

        //--A4 -> B4: 23,45, target cell is empty
        res = proc.CopyCellValue(excelSheet, "A4", excelSheetDest, "B4");
        Assert.IsTrue(res);

        //--A5 -> B5: 
        res = proc.CopyCellValue(excelSheet, "A5", excelSheetDest, "B5");
        Assert.IsTrue(res);

        //--A6 -> B6: 
        res = proc.CopyCellValue(excelSheet, "A6", excelSheetDest, "B6");
        Assert.IsTrue(res);

        //--A7 -> B7: 
        res = proc.CopyCellValue(excelSheet, "A7", excelSheetDest, "B7");
        Assert.IsTrue(res);

        //--A8 -> B8: 
        res = proc.CopyCellValue(excelSheet, "A8", excelSheetDest, "B8");
        Assert.IsTrue(res);

        //--A9 -> B9: 
        res = proc.CopyCellValue(excelSheet, "A9", excelSheetDest, "B9");
        Assert.IsTrue(res);

        res = proc.CopyCellValue(excelSheet, "A10", excelSheetDest, "B10");
        Assert.IsTrue(res);

        res = proc.CopyCellValue(excelSheet, "A11", excelSheetDest, "B11");
        Assert.IsTrue(res);

        res = proc.CopyCellValue(excelSheet, "A12", excelSheetDest, "B12");
        Assert.IsTrue(res);

        // close the files
        proc.CloseExcelFile(excelFile);
        proc.CloseExcelFile(excelFileDest);

        // then open the destination file and check the value of cell B1
        excelFile = proc.OpenExcelFile(filenameDest);
        excelSheet = proc.GetSheetAt(excelFile, 0);

        //--B2: text
        cell = proc.GetCellAt(excelSheet, "B2");
        Assert.IsNotNull(cell);
        excelCellValue = proc.GetCellValue(excelSheet, cell);
        Assert.AreEqual(ExcelCellType.String, excelCellValue.CellType);
        Assert.AreEqual("text", excelCellValue.StringValue);

        //--B3: 12
        cell = proc.GetCellAt(excelSheet, "B3");
        Assert.IsNotNull(cell);
        excelCellValue = proc.GetCellValue(excelSheet, cell);
        Assert.AreEqual(ExcelCellType.Integer, excelCellValue.CellType);
        Assert.AreEqual(12, excelCellValue.IntegerValue);

        //--B6: hello
        cell = proc.GetCellAt(excelSheet, "B6");
        Assert.IsNotNull(cell);
        excelCellValue = proc.GetCellValue(excelSheet, cell);
        Assert.AreEqual(ExcelCellType.String, excelCellValue.CellType);
        Assert.AreEqual("hello", excelCellValue.StringValue);

        //--B8: hello
        cell = proc.GetCellAt(excelSheet, "B8");
        Assert.IsNotNull(cell);
        excelCellValue = proc.GetCellValue(excelSheet, cell);
        Assert.AreEqual(ExcelCellType.String, excelCellValue.CellType);
        Assert.AreEqual("hello", excelCellValue.StringValue);

        //--B9: 12
        cell = proc.GetCellAt(excelSheet, "B9");
        Assert.IsNotNull(cell);
        excelCellValue = proc.GetCellValue(excelSheet, cell);
        Assert.AreEqual(ExcelCellType.Integer, excelCellValue.CellType);
        Assert.AreEqual(12, excelCellValue.IntegerValue);

        //--B12: 12
        cell = proc.GetCellAt(excelSheet, "B12");
        Assert.IsNotNull(cell);
        excelCellValue = proc.GetCellValue(excelSheet, cell);
        Assert.AreEqual(ExcelCellType.Integer, excelCellValue.CellType);
        Assert.AreEqual(12, excelCellValue.IntegerValue);

        proc.CloseExcelFile(excelFile);
    }

    /// <summary>
    /// From a excel/sheet source to another excel/sheet destination, copy the value of a cell to another cell.
    /// </summary>
    [TestMethod]
    public void CopyCellValueDate()
    {
        bool res;
        ExcelProcessor proc = new ExcelProcessor();

        string filename = PathFiles + "CopyCellValueDate.xlsx";
        ExcelFile excelFile = proc.OpenExcelFile(filename);
        ExcelSheet excelSheet = proc.GetSheetAt(excelFile, 0);

        string filenameDest = PathFiles + "CopyCellValueDateDest.xlsx";
        ExcelFile excelFileDest = proc.OpenExcelFile(filenameDest);
        ExcelSheet excelSheetDest = proc.GetSheetAt(excelFileDest, 0);

        ExcelCell cell;
        ExcelCellValue excelCellValue;

        //--do action:

        //--A2 -> B2 (is null)
        res = proc.CopyCellValue(excelSheet, "A2", excelSheetDest, "B2");
        Assert.IsTrue(res);

        res = proc.CopyCellValue(excelSheet, "A3", excelSheetDest, "B3");
        Assert.IsTrue(res);

        res = proc.CopyCellValue(excelSheet, "A4", excelSheetDest, "B4");
        Assert.IsTrue(res);

        res = proc.CopyCellValue(excelSheet, "A5", excelSheetDest, "B5");
        Assert.IsTrue(res);

        res = proc.CopyCellValue(excelSheet, "A6", excelSheetDest, "B6");
        Assert.IsTrue(res);

        // close the files
        proc.CloseExcelFile(excelFile);
        proc.CloseExcelFile(excelFileDest);

        // then open the destination file and check the value of cell B1
        excelFile = proc.OpenExcelFile(filenameDest);
        excelSheet = proc.GetSheetAt(excelFile, 0);

        cell = proc.GetCellAt(excelSheet, "B2");
        Assert.IsNotNull(cell);
        excelCellValue = proc.GetCellValue(excelSheet, cell);
        Assert.AreEqual(ExcelCellType.DateOnly, excelCellValue.CellType);
        Assert.AreEqual(new DateOnly(2026,12,10), excelCellValue.DateOnlyValue);

        cell = proc.GetCellAt(excelSheet, "B3");
        Assert.IsNotNull(cell);
        excelCellValue = proc.GetCellValue(excelSheet, cell);
        Assert.AreEqual(ExcelCellType.DateOnly, excelCellValue.CellType);
        Assert.AreEqual(new DateOnly(2026, 12, 10), excelCellValue.DateOnlyValue);

        cell = proc.GetCellAt(excelSheet, "B4");
        Assert.IsNotNull(cell);
        excelCellValue = proc.GetCellValue(excelSheet, cell);
        Assert.AreEqual(ExcelCellType.DateOnly, excelCellValue.CellType);
        Assert.AreEqual(new DateOnly(2026, 12, 10), excelCellValue.DateOnlyValue);

        cell = proc.GetCellAt(excelSheet, "B5");
        Assert.IsNotNull(cell);
        excelCellValue = proc.GetCellValue(excelSheet, cell);
        Assert.AreEqual(ExcelCellType.DateOnly, excelCellValue.CellType);
        Assert.AreEqual(new DateOnly(2026, 12, 10), excelCellValue.DateOnlyValue);

        cell = proc.GetCellAt(excelSheet, "B6");
        Assert.IsNotNull(cell);
        excelCellValue = proc.GetCellValue(excelSheet, cell);
        Assert.AreEqual(ExcelCellType.DateOnly, excelCellValue.CellType);
        Assert.AreEqual(new DateOnly(2026, 12, 10), excelCellValue.DateOnlyValue);

        proc.CloseExcelFile(excelFile);

    }

    [TestMethod]
    public void CopyCellValueNullBlank()
    {
        bool res;
        ExcelProcessor proc = new ExcelProcessor();

        string filename = PathFiles + "CopyCellValueNullBlank.xlsx";
        ExcelFile excelFile = proc.OpenExcelFile(filename);
        ExcelSheet excelSheet = proc.GetSheetAt(excelFile, 0);

        string filenameDest = PathFiles + "CopyCellValueNullBlankDest.xlsx";
        ExcelFile excelFileDest = proc.OpenExcelFile(filenameDest);
        ExcelSheet excelSheetDest = proc.GetSheetAt(excelFileDest, 0);

        ExcelCell cell;
        ExcelCellValue excelCellValue;

        //--do action:

        res = proc.CopyCellValue(excelSheet, "A2", excelSheetDest, "B2");
        Assert.IsTrue(res);

        res = proc.CopyCellValue(excelSheet, "A3", excelSheetDest, "B3");
        Assert.IsTrue(res);

        // close the files
        proc.CloseExcelFile(excelFile);
        proc.CloseExcelFile(excelFileDest);

        // then open the destination file and check the value of cell B1
        excelFile = proc.OpenExcelFile(filenameDest);
        excelSheet = proc.GetSheetAt(excelFile, 0);

        cell = proc.GetCellAt(excelSheet, "B2");
        Assert.IsNull(cell);

        //--B3 has text, set to blank
        cell = proc.GetCellAt(excelSheet, "B3");
        Assert.IsNotNull(cell);
        excelCellValue = proc.GetCellValue(excelSheet, cell);
        Assert.IsTrue(excelCellValue.IsEmpty);

        proc.CloseExcelFile(excelFile);
    }

    [TestMethod]
    public void CopyCellValueStyle()
    {
        bool res;
        ExcelProcessor proc = new ExcelProcessor();

        string filename = PathFiles + "CopyCellValueStyle.xlsx";
        ExcelFile excelFile = proc.OpenExcelFile(filename);
        ExcelSheet excelSheet = proc.GetSheetAt(excelFile, 0);

        string filenameDest = PathFiles + "CopyCellValueStyleDest.xlsx";
        ExcelFile excelFileDest = proc.OpenExcelFile(filenameDest);
        ExcelSheet excelSheetDest = proc.GetSheetAt(excelFileDest, 0);

        ExcelCell cell;
        ExcelCellValue excelCellValue;

        //--do action:

        res = proc.CopyCellValue(excelSheet, "A2", excelSheetDest, "B2");
        Assert.IsTrue(res);

        res = proc.CopyCellValue(excelSheet, "A3", excelSheetDest, "B3");
        Assert.IsTrue(res);

        res = proc.CopyCellValue(excelSheet, "A4", excelSheetDest, "B4");
        Assert.IsTrue(res);

        res = proc.CopyCellValue(excelSheet, "A5", excelSheetDest, "B5");
        Assert.IsTrue(res);

        // close the files
        proc.CloseExcelFile(excelFile);
        proc.CloseExcelFile(excelFileDest);

        // then open the destination file and check the value of cell B1
        excelFile = proc.OpenExcelFile(filenameDest);
        excelSheet = proc.GetSheetAt(excelFile, 0);


        cell = proc.GetCellAt(excelSheet, "B2");
        Assert.IsNotNull(cell);
        excelCellValue = proc.GetCellValue(excelSheet, cell);
        Assert.AreEqual(ExcelCellType.String, excelCellValue.CellType);
        Assert.AreEqual("text", excelCellValue.StringValue);

        cell = proc.GetCellAt(excelSheet, "B3");
        Assert.IsNotNull(cell);
        excelCellValue = proc.GetCellValue(excelSheet, cell);
        Assert.AreEqual(ExcelCellType.Integer, excelCellValue.CellType);
        Assert.AreEqual(12, excelCellValue.IntegerValue);

        cell = proc.GetCellAt(excelSheet, "B4");
        Assert.IsNotNull(cell);
        excelCellValue = proc.GetCellValue(excelSheet, cell);
        Assert.AreEqual(ExcelCellType.Double, excelCellValue.CellType);
        Assert.AreEqual(23.45, excelCellValue.DoubleValue);

        cell = proc.GetCellAt(excelSheet, "B5");
        Assert.IsNotNull(cell);
        excelCellValue = proc.GetCellValue(excelSheet, cell);
        Assert.AreEqual(ExcelCellType.DateOnly, excelCellValue.CellType);
        Assert.AreEqual(new DateOnly(1997, 07, 03), excelCellValue.DateOnlyValue);

        proc.CloseExcelFile(excelFile);
    }

    [TestMethod]
    public void CopyCellValueCuurency()
    {
        bool res;
        ExcelProcessor proc = new ExcelProcessor();

        string filename = PathFiles + "CopyCellValueCurrency.xlsx";
        ExcelFile excelFile = proc.OpenExcelFile(filename);
        ExcelSheet excelSheet = proc.GetSheetAt(excelFile, 0);

        string filenameDest = PathFiles + "CopyCellValueCurrencyDest.xlsx";
        ExcelFile excelFileDest = proc.OpenExcelFile(filenameDest);
        ExcelSheet excelSheetDest = proc.GetSheetAt(excelFileDest, 0);

        ExcelCell cell;
        ExcelCellValue excelCellValue;

        //--do action:

        res = proc.CopyCellValue(excelSheet, "A2", excelSheetDest, "B2");
        Assert.IsTrue(res);

        res = proc.CopyCellValue(excelSheet, "A3", excelSheetDest, "B3");
        Assert.IsTrue(res);

        // 27€
        res = proc.CopyCellValue(excelSheet, "A4", excelSheetDest, "B4");
        Assert.IsTrue(res);

        // 54 €
        res = proc.CopyCellValue(excelSheet, "A5", excelSheetDest, "B5");
        Assert.IsTrue(res);

        res = proc.CopyCellValue(excelSheet, "A6", excelSheetDest, "B6");
        Assert.IsTrue(res);

        res = proc.CopyCellValue(excelSheet, "A7", excelSheetDest, "B7");
        Assert.IsTrue(res);

        // close the files
        proc.CloseExcelFile(excelFile);
        proc.CloseExcelFile(excelFileDest);

        // then open the destination file and check the value of cell B1
        excelFile = proc.OpenExcelFile(filenameDest);
        excelSheet = proc.GetSheetAt(excelFile, 0);


        cell = proc.GetCellAt(excelSheet, "B2");
        Assert.IsNotNull(cell);
        excelCellValue = proc.GetCellValue(excelSheet, cell);
        Assert.AreEqual(ExcelCellType.String, excelCellValue.CellType);
        Assert.AreEqual("hello", excelCellValue.StringValue);

        cell = proc.GetCellAt(excelSheet, "B3");
        Assert.IsNotNull(cell);
        excelCellValue = proc.GetCellValue(excelSheet, cell);
        Assert.AreEqual(ExcelCellType.Integer, excelCellValue.CellType);
        Assert.AreEqual(12, excelCellValue.IntegerValue);

        // 27 €
        cell = proc.GetCellAt(excelSheet, "B4");
        Assert.IsNotNull(cell);
        excelCellValue = proc.GetCellValue(excelSheet, cell);
        Assert.AreEqual(ExcelCellType.Double, excelCellValue.CellType);
        Assert.AreEqual(27, excelCellValue.DoubleValue);
        Assert.IsNotNull(excelCellValue.Currency);
        Assert.AreEqual(CurrencyName.Euro, excelCellValue.Currency.Name);
        Assert.AreEqual(CurrencyCode.EUR, excelCellValue.Currency.Code);

        // 56 €
        cell = proc.GetCellAt(excelSheet, "B5");
        Assert.IsNotNull(cell);
        excelCellValue = proc.GetCellValue(excelSheet, cell);
        Assert.AreEqual(ExcelCellType.Double, excelCellValue.CellType);
        Assert.AreEqual(56, excelCellValue.DoubleValue);
        Assert.IsNotNull(excelCellValue.Currency);
        Assert.AreEqual(CurrencyName.Euro, excelCellValue.Currency.Name);
        Assert.AreEqual(CurrencyCode.EUR, excelCellValue.Currency.Code);

        // $ 123,45  - US dollar
        cell = proc.GetCellAt(excelSheet, "B6");
        Assert.IsNotNull(cell);
        excelCellValue = proc.GetCellValue(excelSheet, cell);
        Assert.AreEqual(ExcelCellType.Double, excelCellValue.CellType);
        Assert.AreEqual(123.45, excelCellValue.DoubleValue);
        Assert.IsNotNull(excelCellValue.Currency);
        Assert.AreEqual(CurrencyName.UsDollar, excelCellValue.Currency.Name);
        Assert.AreEqual(CurrencyCode.USD, excelCellValue.Currency.Code);

        //  $325,48 - US dollar
        cell = proc.GetCellAt(excelSheet, "B7");
        Assert.IsNotNull(cell);
        excelCellValue = proc.GetCellValue(excelSheet, cell);
        Assert.AreEqual(ExcelCellType.Double, excelCellValue.CellType);
        Assert.AreEqual(325.48, excelCellValue.DoubleValue);
        Assert.IsNotNull(excelCellValue.Currency);
        Assert.AreEqual(CurrencyName.UsDollar, excelCellValue.Currency.Name);
        Assert.AreEqual(CurrencyCode.USD, excelCellValue.Currency.Code);

        proc.CloseExcelFile(excelFile);
    }

}
