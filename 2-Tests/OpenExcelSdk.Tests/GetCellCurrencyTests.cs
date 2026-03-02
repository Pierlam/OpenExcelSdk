using DocumentFormat.OpenXml.Drawing;
using OpenExcelSdk.System;
using OpenExcelSdk.Tests._50_Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcelSdk.Tests;

[TestClass]
public class GetCellCurrencyTests : TestBase
{
    [TestMethod]
    public void GetCellCurrency()
    {
        bool res;
        ExcelProcessor proc = new ExcelProcessor();

        string filename = PathFiles + "GetCellCurrency.xlsx";
        ExcelFile excelFile = proc.OpenExcelFile(filename);

        ExcelSheet excelSheet = proc.GetFirstSheet(excelFile);

        ExcelCell cell;
        ExcelCellValue cellValue;

        //--B2: euro
        cellValue = proc.GetCellValue(excelSheet, "B2");
        Assert.IsNotNull(cellValue);
        Assert.AreEqual(120, cellValue.DoubleValue);
        Assert.IsNotNull(cellValue.Currency);
        Assert.AreEqual(CurrencyName.Euro, cellValue.Currency.Name);
        Assert.AreEqual(CurrencyCode.EUR, cellValue.Currency.Code);

        //--B3: euro
        cellValue = proc.GetCellValue(excelSheet, "B3");
        Assert.IsNotNull(cellValue);
        Assert.AreEqual(345, cellValue.DoubleValue);
        Assert.IsNotNull(cellValue.Currency);
        Assert.AreEqual(CurrencyName.Euro, cellValue.Currency.Name);

        //--B4: US dollar
        cellValue = proc.GetCellValue(excelSheet, "B4");
        Assert.IsNotNull(cellValue);
        Assert.IsNotNull(cellValue.Currency);
        Assert.AreEqual(CurrencyName.UsDollar, cellValue.Currency.Name);
        Assert.AreEqual(CurrencyCode.USD, cellValue.Currency.Code);

        //--B5: US dollar
        cellValue = proc.GetCellValue(excelSheet, "B5");
        Assert.IsNotNull(cellValue);
        Assert.IsNotNull(cellValue.Currency);
        Assert.AreEqual(CurrencyName.UsDollar, cellValue.Currency.Name);
        Assert.AreEqual(CurrencyCode.USD, cellValue.Currency.Code);

        //--B6: bitcoin
        cellValue = proc.GetCellValue(excelSheet, "B6");
        Assert.IsNotNull(cellValue);
        Assert.IsNotNull(cellValue.Currency);
        Assert.AreEqual(CurrencyName.Bitcoin, cellValue.Currency.Name);
        Assert.AreEqual(CurrencyCode.BTC, cellValue.Currency.Code);

        //--B7: bitcoin
        cellValue = proc.GetCellValue(excelSheet, "B7");
        Assert.IsNotNull(cellValue);
        Assert.IsNotNull(cellValue.Currency);
        Assert.AreEqual(CurrencyName.Bitcoin, cellValue.Currency.Name);
        Assert.AreEqual(CurrencyCode.BTC, cellValue.Currency.Code);

        //--B8: australian dollar
        cellValue = proc.GetCellValue(excelSheet, "B8");
        Assert.IsNotNull(cellValue);
        Assert.IsNotNull(cellValue.Currency);
        Assert.AreEqual(CurrencyName.AustralianDollar, cellValue.Currency.Name);
        Assert.AreEqual(CurrencyCode.AUD, cellValue.Currency.Code);

        //--B9: australian dollar
        cellValue = proc.GetCellValue(excelSheet, "B8");
        Assert.IsNotNull(cellValue);
        Assert.IsNotNull(cellValue.Currency);
        Assert.AreEqual(CurrencyName.AustralianDollar, cellValue.Currency.Name);
        Assert.AreEqual(CurrencyCode.AUD, cellValue.Currency.Code);

        //--B10: chinese yuan
        cellValue = proc.GetCellValue(excelSheet, "B10");
        Assert.IsNotNull(cellValue);
        Assert.IsNotNull(cellValue.Currency);
        Assert.AreEqual(CurrencyName.ChineseYuan, cellValue.Currency.Name);
        Assert.AreEqual(CurrencyCode.CNY, cellValue.Currency.Code);

        //--B11: chinese yuan
        cellValue = proc.GetCellValue(excelSheet, "B11");
        Assert.IsNotNull(cellValue);
        Assert.IsNotNull(cellValue.Currency);
        Assert.AreEqual(CurrencyName.ChineseYuan, cellValue.Currency.Name);
        Assert.AreEqual(CurrencyCode.CNY, cellValue.Currency.Code);

        //--B12: canadian dollar
        cellValue = proc.GetCellValue(excelSheet, "B12");
        Assert.IsNotNull(cellValue);
        Assert.IsNotNull(cellValue.Currency);
        Assert.AreEqual(CurrencyName.CanadianDollar, cellValue.Currency.Name);
        Assert.AreEqual(CurrencyCode.CAD, cellValue.Currency.Code);

        //--B13: canadian dollar
        cellValue = proc.GetCellValue(excelSheet, "B13");
        Assert.IsNotNull(cellValue);
        Assert.IsNotNull(cellValue.Currency);
        Assert.AreEqual(CurrencyName.CanadianDollar, cellValue.Currency.Name);
        Assert.AreEqual(CurrencyCode.CAD, cellValue.Currency.Code);

        //--B14: british pound
        cellValue = proc.GetCellValue(excelSheet, "B14");
        Assert.IsNotNull(cellValue);
        Assert.IsNotNull(cellValue.Currency);
        Assert.AreEqual(CurrencyName.BritishPound, cellValue.Currency.Name);
        Assert.AreEqual(CurrencyCode.GBP, cellValue.Currency.Code);

        //--B15: swiss franc
        cellValue = proc.GetCellValue(excelSheet, "B15");
        Assert.IsNotNull(cellValue);
        Assert.IsNotNull(cellValue.Currency);
        Assert.AreEqual(CurrencyName.SwissFranc, cellValue.Currency.Name);
        Assert.AreEqual(CurrencyCode.CHF, cellValue.Currency.Code);

        //--B16: japonese yen
        cellValue = proc.GetCellValue(excelSheet, "B16");
        Assert.IsNotNull(cellValue);
        Assert.IsNotNull(cellValue.Currency);
        Assert.AreEqual(CurrencyName.JapaneseYen, cellValue.Currency.Name);
        Assert.AreEqual(CurrencyCode.JPY, cellValue.Currency.Code);

        //--B17: korean wan
        cellValue = proc.GetCellValue(excelSheet, "B17");
        Assert.IsNotNull(cellValue);
        Assert.IsNotNull(cellValue.Currency);
        Assert.AreEqual(CurrencyName.SouthKoreanWon, cellValue.Currency.Name);
        Assert.AreEqual(CurrencyCode.KWR, cellValue.Currency.Code);
    }
}
