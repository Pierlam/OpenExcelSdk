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
public class SetCellValueCurrencyTests : TestBase
{
    [TestMethod]
    public void SetCellValueCurrency_DEBUG()
    {
        bool res;
        ExcelProcessor proc = new ExcelProcessor();

        string filename = PathFiles + "SetCellValueCurrency_DEBUG.xlsx";
        ExcelFile excelFile = proc.OpenExcelFile(filename);

        ExcelSheet excelSheet = proc.GetSheetAt(excelFile, 0);

        ExcelCell cell;
        ExcelCellValue excelCellValue;

        // to check style/CellFormat creation
        var stylesPart = excelSheet.ExcelFile.WorkbookPart.WorkbookStylesPart;
        int count = stylesPart.Stylesheet.CellFormats.Elements().Count();

        //res = proc.SetCellValueCurrency(excelSheet, "B2", -12.34, CurrencyFormat.Accounting, CurrencyName.Euro, 2);

        //==: US Dollar:
        //--B7: $12,34  with 2 decimals, accounting format
        res = proc.SetCellValueCurrency(excelSheet, "B7", 12.34, CurrencyFormat.Accounting, CurrencyName.UsDollar, 2);
        Assert.IsTrue(res);

        // save the changes
        proc.CloseExcelFile(excelFile);

    }

    [TestMethod]
    public void SetCellValueCurrency()
    {
        bool res;
        ExcelProcessor proc = new ExcelProcessor();

        string filename = PathFiles + "SetCellValueCurrency.xlsx";
        ExcelFile excelFile = proc.OpenExcelFile(filename);

        ExcelSheet excelSheet = proc.GetSheetAt(excelFile, 0);

        ExcelCell cell;
        ExcelCellValue excelCellValue;

        // to check style/CellFormat creation
        var stylesPart = excelSheet.ExcelFile.WorkbookPart.WorkbookStylesPart;
        int count = stylesPart.Stylesheet.CellFormats.Elements().Count();

        //==EURO:
        //--B2: 12,34 € with 2 decimals, accounting format
        res = proc.SetCellValueCurrency(excelSheet, "B2", -12.34, CurrencyFormat.Accounting, CurrencyName.Euro, 2);
        Assert.IsTrue(res);

        //--B3: -392,78 €, 2 decimals, currency format  
        res = proc.SetCellValueCurrency(excelSheet, "B3", -392.78, CurrencyFormat.Currency, CurrencyName.Euro, 2);
        Assert.IsTrue(res);

        //--B4: -550,00 €
        res = proc.SetCellValueCurrency(excelSheet, "B4", -550, CurrencyFormat.CurrencyLeftSpace, CurrencyName.Euro, 2);
        Assert.IsTrue(res);

        //--B5: -71,00 €
        res = proc.SetCellValueCurrency(excelSheet, "B5", -71, CurrencyFormat.CurrencyRedNegativeNoSign, CurrencyName.Euro, 2);
        Assert.IsTrue(res);

        //--B6: 
        res = proc.SetCellValueCurrency(excelSheet, "B6", -62, CurrencyFormat.CurrencyRedNegative, CurrencyName.Euro, 2);
        Assert.IsTrue(res);

        //==: US Dollar:  ->PB HERE
        //--B7: $12,34  with 2 decimals, accounting format
        res = proc.SetCellValueCurrency(excelSheet, "B7", -12.34, CurrencyFormat.Accounting, CurrencyName.UsDollar, 2);
        Assert.IsTrue(res);

        res = proc.SetCellValueCurrency(excelSheet, "B8", -392.78, CurrencyFormat.Currency, CurrencyName.UsDollar, 2);
        Assert.IsTrue(res);

        res = proc.SetCellValueCurrency(excelSheet, "B9", -550, CurrencyFormat.CurrencyLeftSpace, CurrencyName.UsDollar, 2);
        Assert.IsTrue(res);

        res = proc.SetCellValueCurrency(excelSheet, "B10", -71, CurrencyFormat.CurrencyRedNegativeNoSign, CurrencyName.UsDollar, 2);
        Assert.IsTrue(res);

        res = proc.SetCellValueCurrency(excelSheet, "B11", -62, CurrencyFormat.CurrencyRedNegative, CurrencyName.UsDollar, 2);
        Assert.IsTrue(res);

        //==: Swiss Franc/CHF


        // save the changes
        proc.CloseExcelFile(excelFile);

        //==>check the excel content
        excelFile = proc.OpenExcelFile(filename);
        excelSheet = proc.GetSheetAt(excelFile, 0);


        //==EURO:
        //--B2: 123,00 € with 2 decimals, accounting format
        cell = proc.GetCellAt(excelSheet, "B2");
        Assert.IsNotNull(cell);
        excelCellValue = proc.GetCellValue(excelSheet, cell);
        Assert.AreEqual(ExcelCellType.Double, excelCellValue.CellType);
        Assert.AreEqual(-12.34, excelCellValue.DoubleValue);
        Assert.IsNotNull(excelCellValue.Currency);
        Assert.AreEqual("€", excelCellValue.Currency.Symbol);
        Assert.AreEqual(CurrencyName.Euro, excelCellValue.Currency.Name);
        Assert.AreEqual(CurrencyCode.EUR, excelCellValue.Currency.Code);
        Assert.AreEqual("\"€\"", excelCellValue.Currency.ExcelCode);
        Assert.IsTrue(excelCellValue.NumberFormat.Contains("#0.00"));
        Assert.AreEqual(CurrencyFormat.Accounting, excelCellValue.Currency.Format);


        //--B3: -392,78 €, 2 decimals, currency format  
        cell = proc.GetCellAt(excelSheet, "B3");
        Assert.IsNotNull(cell);
        excelCellValue = proc.GetCellValue(excelSheet, cell);
        Assert.AreEqual(ExcelCellType.Double, excelCellValue.CellType);
        Assert.AreEqual(-392.78, excelCellValue.DoubleValue);
        Assert.IsNotNull(excelCellValue.Currency);
        Assert.AreEqual("€", excelCellValue.Currency.Symbol);
        Assert.AreEqual(CurrencyName.Euro, excelCellValue.Currency.Name);
        Assert.AreEqual(CurrencyCode.EUR, excelCellValue.Currency.Code);
        Assert.AreEqual("\"€\"", excelCellValue.Currency.ExcelCode);
        Assert.IsTrue(excelCellValue.NumberFormat.Contains("#0.00"));
        Assert.AreEqual(CurrencyFormat.Currency, excelCellValue.Currency.Format);

        //--B4: -550 €
        cell = proc.GetCellAt(excelSheet, "B4");
        Assert.IsNotNull(cell);
        excelCellValue = proc.GetCellValue(excelSheet, cell);
        Assert.AreEqual(ExcelCellType.Double, excelCellValue.CellType);
        Assert.AreEqual(-550, excelCellValue.DoubleValue);
        Assert.IsNotNull(excelCellValue.Currency);
        Assert.AreEqual("€", excelCellValue.Currency.Symbol);
        Assert.AreEqual(CurrencyName.Euro, excelCellValue.Currency.Name);
        Assert.AreEqual(CurrencyCode.EUR, excelCellValue.Currency.Code);
        Assert.AreEqual("\"€\"", excelCellValue.Currency.ExcelCode);
        Assert.IsTrue(excelCellValue.NumberFormat.Contains("#0.00"));
        Assert.AreEqual(CurrencyFormat.CurrencyLeftSpace, excelCellValue.Currency.Format);

        //--B5: 
        cell = proc.GetCellAt(excelSheet, "B5");
        Assert.IsNotNull(cell);
        excelCellValue = proc.GetCellValue(excelSheet, cell);
        Assert.AreEqual(ExcelCellType.Double, excelCellValue.CellType);
        Assert.AreEqual(-71, excelCellValue.DoubleValue);
        Assert.IsNotNull(excelCellValue.Currency);
        Assert.AreEqual("€", excelCellValue.Currency.Symbol);
        Assert.AreEqual(CurrencyName.Euro, excelCellValue.Currency.Name);
        Assert.AreEqual(CurrencyCode.EUR, excelCellValue.Currency.Code);
        Assert.AreEqual("\"€\"", excelCellValue.Currency.ExcelCode);
        Assert.IsTrue(excelCellValue.NumberFormat.Contains("#0.00"));
        Assert.AreEqual(CurrencyFormat.CurrencyRedNegativeNoSign, excelCellValue.Currency.Format);

        //--B5: 
        cell = proc.GetCellAt(excelSheet, "B6");
        Assert.IsNotNull(cell);
        excelCellValue = proc.GetCellValue(excelSheet, cell);
        Assert.AreEqual(ExcelCellType.Double, excelCellValue.CellType);
        Assert.AreEqual(-62, excelCellValue.DoubleValue);
        Assert.IsNotNull(excelCellValue.Currency);
        Assert.AreEqual("€", excelCellValue.Currency.Symbol);
        Assert.AreEqual(CurrencyName.Euro, excelCellValue.Currency.Name);
        Assert.AreEqual(CurrencyCode.EUR, excelCellValue.Currency.Code);
        Assert.AreEqual("\"€\"", excelCellValue.Currency.ExcelCode);
        Assert.IsTrue(excelCellValue.NumberFormat.Contains("#0.00"));
        Assert.AreEqual(CurrencyFormat.CurrencyRedNegative, excelCellValue.Currency.Format);


        //==: US Dollar:
        cell = proc.GetCellAt(excelSheet, "B7");
        Assert.IsNotNull(cell);
        excelCellValue = proc.GetCellValue(excelSheet, cell);
        Assert.AreEqual(ExcelCellType.Double, excelCellValue.CellType);
        Assert.AreEqual(-12.34, excelCellValue.DoubleValue);
        Assert.IsNotNull(excelCellValue.Currency);
        Assert.AreEqual("$", excelCellValue.Currency.Symbol);
        Assert.AreEqual(CurrencyName.UsDollar, excelCellValue.Currency.Name);
        Assert.AreEqual(CurrencyCode.USD, excelCellValue.Currency.Code);
        Assert.AreEqual("[$$-409]", excelCellValue.Currency.ExcelCode);
        Assert.IsTrue(excelCellValue.NumberFormat.Contains("#0.00"));
        Assert.AreEqual(CurrencyFormat.Accounting, excelCellValue.Currency.Format);


        //==: Swiss Franc/CHF

        // todo: add others tests



        proc.CloseExcelFile(excelFile);

        //Assert.Fail("Test not implemented yet");
    }

}
