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
    public static Currency CreateEuro(string numberFormat)
    {
        return new Currency()
        {
            Symbol = "€",
            Code = CurrencyCode.EUR,
            ExcelCode = "\"€\"",
            Name = CurrencyName.Euro,
            SymbolPosition = CurrencySymbolPosition.After,
            Format = GetCurrencyFormat(numberFormat)
        };
    }

    public static Currency CreateUsDollar(string numberFormat)
    {
        return new Currency()
        {
            Symbol = "$",
            Code = CurrencyCode.USD,
            ExcelCode = "[$$-409]",
            Name = CurrencyName.UsDollar,
            SymbolPosition = CurrencySymbolPosition.Before,
            Format = GetCurrencyFormat(numberFormat)
        };
    }

    public static Currency CreateJapaneseYen(string numberFormat)
    {
        return new Currency()
        {
            Symbol = "¥",
            Code = CurrencyCode.JPY,
            ExcelCode = "[$¥-411]",
            Name = CurrencyName.JapaneseYen,
            SymbolPosition = CurrencySymbolPosition.Before,
            Format = GetCurrencyFormat(numberFormat)
        };
    }

    public static Currency CreateSouthKoreanWon(string numberFormat)
    { 
        return new Currency()
        {
            Symbol = "₩",
            Code = CurrencyCode.KWR,
            ExcelCode = "[$₩-412]",
            Name = CurrencyName.SouthKoreanWon,
            SymbolPosition = CurrencySymbolPosition.Before,
            Format = GetCurrencyFormat(numberFormat)
        };
    }

    /// <summary>
    /// Creates a new instance of the Swiss Franc currency configured with the specified number format.
    /// can be [$CHF-417] or also:
    ///  #,##0.00\\ [$CHF]"
    /// </summary>
    /// <remarks>The returned Currency instance uses the Swiss Franc symbol and Excel code, and places the
    /// symbol after the numeric value. Ensure that the provided number format string is valid to avoid display
    /// issues.</remarks>
    /// <param name="numberFormat">The format string that determines how currency values are displayed. Must be a valid Excel-compatible format.</param>
    /// <returns>A Currency object representing the Swiss Franc, initialized with its standard symbol, code, and formatting.</returns>
    public static Currency CreateSwissFranc(string numberFormat)
    { 
        return new Currency()
        {
            Symbol = "CHF",
            Code = CurrencyCode.CHF,
            ExcelCode = "[$CHF-417]",
            Name = CurrencyName.SwissFranc,
            SymbolPosition = CurrencySymbolPosition.After,
            Format = GetCurrencyFormat(numberFormat)
        };
    }

    public static Currency CreateChineseYuan(string numberFormat)
    {
        return new Currency()
        {
            Symbol = "¥",
            Code = CurrencyCode.CNY,
            ExcelCode = "[$¥-804]",
            Name = CurrencyName.ChineseYuan,
            SymbolPosition = CurrencySymbolPosition.Before,
            Format = GetCurrencyFormat(numberFormat)
        };
    }

    public static Currency CreateBritishPound(string numberFormat)
    {
        return new Currency()
        {
            Symbol = "£",
            Code = CurrencyCode.GBP,
            ExcelCode = "[$£-809]",
            Name = CurrencyName.BritishPound,
            SymbolPosition = CurrencySymbolPosition.Before,
            Format = GetCurrencyFormat(numberFormat)
        };
    }

    public static Currency CreateAustralianDollar(string numberFormat)
    {
        return new Currency()
        {
            Symbol = "$",
            Code = CurrencyCode.AUD,
            ExcelCode = "[$$-C09]",
            Name = CurrencyName.AustralianDollar,
            SymbolPosition = CurrencySymbolPosition.Before,
            Format = GetCurrencyFormat(numberFormat)
        };
    }

    public static Currency CreateCanadianDollar(string numberFormat)
    {
        return new Currency()
        {
            Symbol = "$",
            Code = CurrencyCode.CAD,
            ExcelCode = "[$$-1009]",
            Name = CurrencyName.CanadianDollar,
            SymbolPosition = CurrencySymbolPosition.Before,
            Format = GetCurrencyFormat(numberFormat)
        };
    }

    public static Currency CreateNewZealandDollar(string numberFormat)
    {
        return new Currency()
        {
            Symbol = "$",
            Code = CurrencyCode.NZD,
            ExcelCode = "[$$-1409]",
            Name = CurrencyName.NewZealandDollar,
            SymbolPosition = CurrencySymbolPosition.Before,
            Format = GetCurrencyFormat(numberFormat)
        };
    }

    public static Currency CreateSingaporeDollar(string numberFormat)
    {
        return new Currency()
        {
            Symbol = "$",
            Code = CurrencyCode.SGD,
            ExcelCode = "[$$-4809]",
            Name = CurrencyName.SingaporeDollar,
            SymbolPosition = CurrencySymbolPosition.Before,
            Format = GetCurrencyFormat(numberFormat)
        };
    }

    public static Currency CreateBitcoin(string numberFormat)
    {
        return new Currency()
        {
            Symbol = "₿",
            Code = CurrencyCode.BTC,
            ExcelCode = "[$₿]",
            Name = CurrencyName.Bitcoin,
            SymbolPosition = CurrencySymbolPosition.Before,
            Format = GetCurrencyFormat(numberFormat)
        };
    }

    public static Currency CreateNotDefined(string numberFormat)
    {
        return new Currency()
        {
            Symbol = "?",
            Code = CurrencyCode.XXX,
            ExcelCode = "\"?\"",
            Name = CurrencyName.NotDefined,
            SymbolPosition = CurrencySymbolPosition.Before,
            Format = GetCurrencyFormat(numberFormat)
        };
    }

    /// <summary>
    /// Return the currency format based on the number format provided. If the number format is not recognized, it will return the default currency format.
    /// </summary>
    /// <param name="numberFormat"></param>
    /// <returns></returns>
    public static CurrencyFormat GetCurrencyFormat(string numberFormat)
    {
        if(string.IsNullOrWhiteSpace(numberFormat))
            // return the default value
            return CurrencyFormat.Currency;

        //--currency format, Negative in red: #,##0.00\ "€";[Red]\-#,##0.00\ "€" 
        if (numberFormat.Contains("[Red]") && numberFormat.Contains("\\-"))
            return CurrencyFormat.CurrencyRedNegative;

        //--currency format, Negative in red, no negative sign: #,##0.00\ "€";[Red]#,##0.00\ "€" -or- [$$-409]#,##0.00;[Red][$$-409]#,##0.00
        if (numberFormat.Contains("[Red]"))
            return CurrencyFormat.CurrencyRedNegativeNoSign;

        // accounting format: _-* #,##0.00\ "€"_-;\-* #,##0.00\ "€"_-;_-* "-"??\ "€"_-;_-@_-
        if (numberFormat.Contains("??"))
            return CurrencyFormat.Accounting;

        //--Currency format, space on left side:  "#,##0.00\\ \"€\";\\-#,##0.00\\ \"€\"" -or- [$$-409]#,##0.00_ ;\-[$$-409]#,##0.00\
        if (numberFormat.Contains("\\-"))
            return CurrencyFormat.CurrencyLeftSpace;

        // default currency format: #,##0.00\ "€"
        return CurrencyFormat.Currency;
    }
}
