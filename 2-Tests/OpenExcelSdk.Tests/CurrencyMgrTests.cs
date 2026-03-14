using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using OpenExcelSdk.System;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcelSdk.Tests;

[TestClass]
public class CurrencyMgrTests
{
    [TestMethod]
    public void SetCellValueCurrencyEuro()
    {
        bool res = CurrencyMgr.CreateNumberFormat(CurrencyFormat.Currency, CurrencyName.Euro, 0, out string numberFormat);
        Assert.IsTrue(res);
        Assert.AreEqual("#,##0\\ \"€\"", numberFormat);

        res = CurrencyMgr.CreateNumberFormat(CurrencyFormat.Currency, CurrencyName.Euro, 1, out numberFormat);
        Assert.IsTrue(res);
        Assert.AreEqual("#,##0.0\\ \"€\"", numberFormat);

        res = CurrencyMgr.CreateNumberFormat(CurrencyFormat.Currency, CurrencyName.Euro, 2, out numberFormat);
        Assert.IsTrue(res);
        Assert.AreEqual("#,##0.00\\ \"€\"", numberFormat);

        res = CurrencyMgr.CreateNumberFormat(CurrencyFormat.Currency, CurrencyName.Euro, 4, out numberFormat);
        Assert.IsTrue(res);
        Assert.AreEqual("#,##0.0000\\ \"€\"", numberFormat);
    }

    [TestMethod]
    public void SetCellValueCurrencyUsDollar()
    {
        bool res = CurrencyMgr.CreateNumberFormat(CurrencyFormat.Currency, CurrencyName.UsDollar, 0, out string numberFormat);
        Assert.IsTrue(res);
        Assert.AreEqual("[$$-409]#,##0", numberFormat);

        res = CurrencyMgr.CreateNumberFormat(CurrencyFormat.Currency, CurrencyName.UsDollar, 1, out numberFormat);
        Assert.IsTrue(res);
        Assert.AreEqual("[$$-409]#,##0.0", numberFormat);

        res = CurrencyMgr.CreateNumberFormat(CurrencyFormat.Currency, CurrencyName.UsDollar, 2, out numberFormat);
        Assert.IsTrue(res);
        Assert.AreEqual("[$$-409]#,##0.00", numberFormat);

        res = CurrencyMgr.CreateNumberFormat(CurrencyFormat.Currency, CurrencyName.UsDollar, 4, out numberFormat);
        Assert.IsTrue(res);
        Assert.AreEqual("[$$-409]#,##0.0000", numberFormat);
    }

    [TestMethod]
    public void SetCellValueCurrencyBitcoin()
    {
        bool res = CurrencyMgr.CreateNumberFormat(CurrencyFormat.Currency, CurrencyName.Bitcoin, 6, out string numberFormat);
        Assert.IsTrue(res);
        Assert.AreEqual("[$₿]#,##0.000000", numberFormat);
    }

    [TestMethod]
    public void SetCellValueCurrencySwissFranc()
    {
        bool res = CurrencyMgr.CreateNumberFormat(CurrencyFormat.Currency, CurrencyName.SwissFranc, 2, out string numberFormat);
        Assert.IsTrue(res);
        Assert.AreEqual("#,##0.00\\ [$CHF-417]", numberFormat);
    }

    /// <summary>
    /// _-* #,##0.00\ "€"_-;\-* #,##0.00\ "€"_-;_-* "-"??\ "€"_-;_-@_-
    /// Don't forget the last char! which is a space char! 
    /// </summary>
    [TestMethod]
    public void SetCellValueAccountingEuro()
    {
        bool res = CurrencyMgr.CreateNumberFormat(CurrencyFormat.Accounting, CurrencyName.Euro, 0, out string numberFormat);
        Assert.IsTrue(res);
        Assert.AreEqual("_-* #,##0\\ \"€\"_-;\\-* #,##0\\ \"€\"_-;_-* \"-\"??\\ \"€\"_-;_-@_- ", numberFormat);

        res = CurrencyMgr.CreateNumberFormat(CurrencyFormat.Accounting, CurrencyName.Euro, 2, out numberFormat);
        Assert.IsTrue(res);
        Assert.AreEqual("_-* #,##0.00\\ \"€\"_-;\\-* #,##0.00\\ \"€\"_-;_-* \"-\"??\\ \"€\"_-;_-@_- ", numberFormat);
    }

    /// <summary>
    /// Don't forget the last char! which is a space char! 
    /// </summary>
    [TestMethod]
    public void SetCellValueAccountingUsDollar()
    {
        bool res = CurrencyMgr.CreateNumberFormat(CurrencyFormat.Accounting, CurrencyName.UsDollar, 0, out string numberFormat);
        Assert.IsTrue(res);
        Assert.AreEqual("_-[$$-409]* #,##0_ ;_-[$$-409]* \\-#,##0\\ ;_-[$$-409]* \"-\"??_ ;_-@_ ", numberFormat);

        res = CurrencyMgr.CreateNumberFormat(CurrencyFormat.Accounting, CurrencyName.UsDollar, 2, out numberFormat);
        Assert.IsTrue(res);
        Assert.AreEqual("_-[$$-409]* #,##0.00_ ;_-[$$-409]* \\-#,##0.00\\ ;_-[$$-409]* \"-\"??_ ;_-@_ ", numberFormat);
    }
}