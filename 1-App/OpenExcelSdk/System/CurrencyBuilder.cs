using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.EMMA;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenExcelSdk.System;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcelSdk.System;

public class CurrencyBuilder
{
    public static Currency CreateEuro()
    {
        return new Currency()
        {
            Symbol = "€",
            Code = CurrencyCode.EUR,
            ExcelCode = "\"€\"",
            Name = CurrencyName.Euro,
            SymbolPosition = CurrencySymbolPosition.After
        };
    }

    public static Currency CreateUsDollar()
    {
        return new Currency()
        {
            Symbol = "$",
            Code = CurrencyCode.USD,
            ExcelCode = "[$$-409]",
            Name = CurrencyName.UsDollar,
            SymbolPosition = CurrencySymbolPosition.Before
        };

    }

    public static Currency CreateJapaneseYen()
    {
        return new Currency()
        {
            Symbol = "¥",
            Code = CurrencyCode.JPY,
            ExcelCode = "[$¥-411]",
            Name = CurrencyName.JapaneseYen,
            SymbolPosition = CurrencySymbolPosition.Before
        };
    }

    public static Currency CreateSouthKoreanWon()
    { 
        return new Currency()
        {
            Symbol = "₩",
            Code = CurrencyCode.KWR,
            ExcelCode = "[$₩-412]",
            Name = CurrencyName.SouthKoreanWon,
            SymbolPosition = CurrencySymbolPosition.Before
        };
    }

    public static Currency CreateSwissFranc()
    { 
        return new Currency()
        {
            Symbol = "CHF",
            Code = CurrencyCode.CHF,
                ExcelCode = "[$CHF-417]",
            Name = CurrencyName.SwissFranc,
            SymbolPosition = CurrencySymbolPosition.After
         };
    }

    public static Currency CreateChineseYuan()
    {
        return new Currency()
        {
            Symbol = "¥",
            Code = CurrencyCode.CNY,
            ExcelCode = "[$¥-804]",
            Name = CurrencyName.ChineseYuan,
            SymbolPosition = CurrencySymbolPosition.Before
        };
    }

    public static Currency CreateBritishPound()
    {
        return new Currency()
        {
            Symbol = "£",
            Code = CurrencyCode.GBP,
            ExcelCode = "[$£-809]",
            Name = CurrencyName.BritishPound,
            SymbolPosition = CurrencySymbolPosition.Before
        };
    }

    public static Currency CreateAustralianDollar()
    {
        return new Currency()
        {
            Symbol = "$",
            Code = CurrencyCode.AUD,
            ExcelCode = "[$$-C09]",
            Name = CurrencyName.AustralianDollar,
            SymbolPosition = CurrencySymbolPosition.Before
        };
    }

    public static Currency CreateCanadianDollar()
    {
        return new Currency()
        {
            Symbol = "$",
            Code = CurrencyCode.CAD,
            ExcelCode = "[$$-1009]",
            Name = CurrencyName.CanadianDollar,
            SymbolPosition = CurrencySymbolPosition.Before
        };
    }

    public static Currency CreateNewZealandDollar()
    {
        return new Currency()
        {
            Symbol = "$",
            Code = CurrencyCode.NZD,
            ExcelCode = "[$$-1409]",
            Name = CurrencyName.NewZealandDollar,
            SymbolPosition = CurrencySymbolPosition.Before
        };
    }

    public static Currency CreateSingaporeDollar()
    {
        return new Currency()
        {
            Symbol = "$",
            Code = CurrencyCode.SGD,
            ExcelCode = "[$$-4809]",
            Name = CurrencyName.SingaporeDollar,
            SymbolPosition = CurrencySymbolPosition.Before
        };
    }

    public static Currency CreateBitcoin()
    {
        return new Currency()
        {
            Symbol = "₿",
            Code = CurrencyCode.BTC,
            ExcelCode = "[$₿]",
            Name = CurrencyName.Bitcoin,
            SymbolPosition = CurrencySymbolPosition.Before
        };
    }

    public static Currency CreateNotDefined()
    {
        return new Currency()
        {
            Symbol = "?",
            Code = CurrencyCode.XXX,
            ExcelCode = "\"?\"",
            Name = CurrencyName.NotDefined,
            SymbolPosition = CurrencySymbolPosition.Before
        };
    }
}
