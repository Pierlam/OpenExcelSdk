using OpenExcelSdk.System;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcelSdk;

public class CurrencyMgr
{
    /// <summary>
    /// 
    /// some cases:
    /// "_(\"$\"* #,##0.00_);_(\"$\"* \\(#,##0.00\\);_(\"$\"* \"-\"??_);_(@_)"
    /// "[$$-409]#,##0.00"
    /// "_-[$\u20bf]\\ * #,##0.000000_-;\\-[$\u20bf]\\ * #,##0.000000_-;_-[$\u20bf]\\ * \"-\"??????_-;_-@_-"
    /// "[$$-C09]#,##0.00"
    /// "[$¥-804]#,##0.00"
    /// </summary>
    /// <param name="numberFormat"></param>
    /// <returns></returns>
    public static Currency CreateCurrency(string numberFormat)
    {
        if (string.IsNullOrEmpty(numberFormat)) return null;

        // check for euro
        if (numberFormat.Contains("\"€") || numberFormat.Contains("[$€-") )
        {
            return new Currency()
            {
                Symbol = "€",
                Code = CurrencyCode.EUR,
                Name = CurrencyName.Euro
            };
        }

        // [$$-1009]#,##0.00 Canadian dollar
        if (numberFormat.Contains("[$$-1009]"))
        {
            return new Currency()
            {
                Symbol = "$",
                Code = CurrencyCode.CAD,
                Name = CurrencyName.CanadianDollar
            };
        }

        // _-[$₿]\ * #,##0.000000_-;\-[$₿]\ * #,##0.000000_-;_-[$₿]\ * "-"??????_-;_-@_-:  bitcoin symbol
        if (numberFormat.Contains("[$₿]"))
        {
            return new Currency()
            {
                Symbol = "₿",
                Code = CurrencyCode.BTC,
                Name = CurrencyName.Bitcoin
            };
        }

        // [$$-C09]#,##0.00: australian dollar
        if (numberFormat.Contains("[$$-C09]"))
        {
            return new Currency()
            {
                Symbol = "$",
                Code = CurrencyCode.AUD,
                Name = CurrencyName.AustralianDollar
            };
        }

        // [$CHF-417] swiss franc
        if (numberFormat.Contains("[$CHF-417]"))
        {
            return new Currency()
            {
                Symbol = "F",
                Code = CurrencyCode.CHF,
                Name = CurrencyName.SwissFranc
            };
        }

        // [$¥-411]#,##0.00 japanese yen
        if (numberFormat.Contains("[$¥-411]"))
        {
            return new Currency()
            {
                Symbol = "¥",
                Code = CurrencyCode.JPY,
                Name = CurrencyName.JapaneseYen
            };
        }

        // [$₩-412]#,##0.00 south korean won
        if (numberFormat.Contains("[$₩-412]"))
        {
            return new Currency()
            {
                Symbol = "₩",
                Code = CurrencyCode.KWR,
                Name = CurrencyName.SouthKoreanWon
            };
        }

        //[$NZ$-409]#,##0.00 new zealand dollar
        if (numberFormat.Contains("[$NZ$-409]"))
        {
            return new Currency()
            {
                Symbol = "$",
                Code = CurrencyCode.NZD,
                Name = CurrencyName.NewZealandDollar
            };
        }

        // [$¥-804]#,##0.00: chinese yuan
        if (numberFormat.Contains("[$¥-804]"))
        {
            return new Currency()
            {
                Symbol = "¥",
                Code = CurrencyCode.CNY,
                Name = CurrencyName.ChineseYuan
            };
        }

        // [$£-809]#,##0.00
        if (numberFormat.Contains("[$£-809]"))
        {
            return new Currency()
            {
                Symbol = "£",
                Code = CurrencyCode.GBP,
                Name = CurrencyName.BritishPound
            };
        }

        // "[$$-409] check for dollar
        if (numberFormat.Contains("[$$-409]"))
        {
            return new Currency()
            {
                Symbol = "$",
                Code = CurrencyCode.USD,
                Name = CurrencyName.UsDollar
            };
        }

        // "[$?-???] not yet managed
        if (numberFormat.Contains("[$"))
        {
            return new Currency()
            {
                Symbol = "?",
                Code = CurrencyCode.XXX,
                Name = CurrencyName.NotDefined
            };
        }

        // add more currencies as needed
        return null;
    }
}
