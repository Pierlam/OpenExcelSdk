using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
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
    /// Create a currency numberformat based on the currency format, currency name and the number of digits after the decimal point. The method constructs a number format string that can be used in Excel to format cells as currency. The format is determined by the specified currency format (e.g., Accounting or Currency), the currency name (e.g., Euro, US Dollar), and the number of digits to display after the decimal point. The resulting number format string is returned through an output parameter.
    /// </summary>
    /// <param name="currencyFormat"></param>
    /// <param name="currencyName"></param>
    /// <param name="digitAfter"></param>
    /// <param name="NumberFormat"></param>
    /// <returns></returns>
    public static bool CreateNumberFormat(CurrencyFormat currencyFormat, CurrencyName currencyName, int digitAfter, out string numberFormat)
    {
        // format: Currency -> exp:  #,##0.00\ "€"
        if (currencyFormat == CurrencyFormat.Currency)
            return CreateNumberFormatCurrency(currencyName, digitAfter, out numberFormat);

        // format: Accounting -> exp: _-* #,##0.00\ "€"_-;\-* #,##0.00\ "€"_-;_-* "-"??\ "€"_-;_-@_-
        if (currencyFormat == CurrencyFormat.Accounting)
            return CreateNumberFormatAccounting(currencyName, digitAfter, out numberFormat);

        numberFormat = null;
        return false;
    }

    /// <summary>
    /// Creates a currency-specific number format string with a defined number of decimal places.
    /// format: Currency -> exp:  #,##0.00\ "€"
    /// </summary>
    /// <remarks>This method returns false if the specified currency is not supported or if the digitAfter
    /// parameter is out of range.</remarks>
    /// <param name="currencyName">The currency for which the number format is generated. Determines the currency symbol and formatting
    /// conventions.</param>
    /// <param name="digitAfter">The number of decimal places to include in the formatted currency string. Must be a non-negative integer.</param>
    /// <param name="numberFormat">When this method returns, contains the resulting number format string for the specified currency.</param>
    /// <returns>true if the number format was successfully created; otherwise, false.</returns>
    public static bool CreateNumberFormatCurrency(CurrencyName currencyName, int digitAfter, out string numberFormat)
    {
        // [$$-409]#,##0.00
        // #,##0.00\ "€"
        // #,##0\ "€"
        numberFormat = "#,##0";

        Currency c = GetCurrency(currencyName);

        // first add digit after decimal point
        if(digitAfter > 0)
        {
            numberFormat += ".";
            for (int i = 0; i < digitAfter; i++)
                numberFormat += "0";
        }

        // then add currency symbol, beofre or after the value
        if(c.SymbolPosition== CurrencySymbolPosition.After)
        {
            // add space char + code e.g 
            numberFormat += "\\ " + c.ExcelCode;
            return true;

        }

        // currency symbol before the value
        numberFormat = c.ExcelCode+numberFormat;
        return true;
    }

    /// <summary>
    /// Creates a number format string based on the specified currency and number of decimal digits.
    /// exp1: _-* #,##0.00\ "€"_-;\-* #,##0.00\ "€"_-;_-* "-"??\ "€"_-;_-@_-
    /// exp2: _-[$$-409]* #,##0.00_ ;_-[$$-409]* \-#,##0.00\ ;_-[$$-409]* "-"??_ ;_-@_
    /// </summary>
    /// <remarks>This method returns false if the provided currency name is not recognized or if digitAfter is
    /// out of the valid range.</remarks>
    /// <param name="currencyName">The currency to use for determining the formatting style. Must be a valid value of the CurrencyName enumeration.</param>
    /// <param name="digitAfter">The number of decimal places to include in the formatted number. Must be a non-negative integer.</param>
    /// <param name="numberFormat">When this method returns, contains the resulting number format string if the operation succeeds; otherwise,
    /// contains an undefined value.</param>
    /// <returns>true if the number format string was successfully created; otherwise, false.</returns>
    public static bool CreateNumberFormatAccounting(CurrencyName currencyName, int digitAfter, out string numberFormat)
    {
        numberFormat = "#,##0";

        Currency c = GetCurrency(currencyName);

        // first add digit after decimal point
        if (digitAfter > 0)
        {
            numberFormat += ".";
            for (int i = 0; i < digitAfter; i++)
                numberFormat += "0";
        }

        // then add currency symbol, after the value
        // exp1: _-* #,##0.00\ "€"_-;\-* #,##0.00\ "€"_-;_-* "-"??\ "€"_-;_-@_-
        if (c.SymbolPosition == CurrencySymbolPosition.After)
        {
            // add space char + code e.g 
            numberFormat = "_-* " + numberFormat + "\\ " +c.ExcelCode + "_-;\\-* " + numberFormat + "\\ " + c.ExcelCode + "_-;_-* \"-\"??\\ " + c.ExcelCode+ "_-;_-@_-";
            return true;
        }

        // currency symbol before the value
        // exp2: _-[$$-409]* #,##0.00_ ;_-[$$-409]* \-#,##0.00\ ;_-[$$-409]* "-"??_ ;_-@_
        numberFormat = "_-" + c.ExcelCode+ "* " + numberFormat+ "_ ;_-" + c.ExcelCode+ "* \\-" + numberFormat+ "\\ ;_-" + c.ExcelCode+ "* \"-\"??_ ;_-@_";
        return true;
    }

    /// <summary>
    /// Create a Currency object based on the number format string. The method checks for specific patterns in the number format string that are commonly used to represent different currencies. If a match is found, it returns a Currency object with the appropriate symbol, code, and name. If no match is found, it returns null. Note that this method may not be exhaustive and may need to be updated to cover additional currencies or formats as needed.
    /// 
    /// Windows Locale Codes - Language are in hexadecimal.
    /// https://www.science.co.il/language/Locale-codes.php
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
            return CurrencyBuilder.CreateEuro(numberFormat);
        }

        // "[$$-409] check for dollar
        if (numberFormat.Contains("[$$-409]"))
        {
            return CurrencyBuilder.CreateUsDollar(numberFormat);
        }

        // [$¥-411]#,##0.00 japanese yen
        if (numberFormat.Contains("[$¥-411]"))
        {
            return CurrencyBuilder.CreateJapaneseYen(numberFormat);
        }

        // [$₩-412]#,##0.00 south korean won
        if (numberFormat.Contains("[$₩-412]"))
        {
            return CurrencyBuilder.CreateSouthKoreanWon(numberFormat);
        }

        // [$CHF-417] or [$CHF] swiss franc
        if (numberFormat.Contains("[$CHF-417]") || numberFormat.Contains("[$CHF]"))
        {
            return CurrencyBuilder.CreateSwissFranc(numberFormat);
        }

        // [$¥-804]#,##0.00: chinese yuan
        if (numberFormat.Contains("[$¥-804]"))
        {
            return CurrencyBuilder.CreateChineseYuan(numberFormat);
        }

        // [$£-809]#,##0.00 british pound
        if (numberFormat.Contains("[$£-809]"))
        {
            return CurrencyBuilder.CreateBritishPound(numberFormat);
        }

        // [$$-C09]#,##0.00: australian dollar
        if (numberFormat.Contains("[$$-C09]"))
        {
            return CurrencyBuilder.CreateAustralianDollar(numberFormat);
        }

        // [$$-1009]#,##0.00 Canadian dollar
        if (numberFormat.Contains("[$$-1009]"))
        {
            return CurrencyBuilder.CreateCanadianDollar(numberFormat);
        }

        //[$$-409]#,##0.00 new zealand dollar
        if (numberFormat.Contains("[$$-1409]"))
        {
            return CurrencyBuilder.CreateNewZealandDollar(numberFormat);
        }

        //[$$-4809]#,##0.00 singapore dollar
        if (numberFormat.Contains("[$$-4809]"))
        {
            return CurrencyBuilder.CreateSingaporeDollar(numberFormat);
        }

        // [$₿]: bitcoin
        if (numberFormat.Contains("[$₿]"))
        {
            return CurrencyBuilder.CreateBitcoin(numberFormat);
        }

        // "[$?-???] not yet managed
        if (numberFormat.Contains("[$"))
        {
            return CurrencyBuilder.CreateNotDefined(numberFormat);
        }

        // add more currencies as needed
        return null;
    }

    /// <summary>
    /// Create currency object by the name.
    /// </summary>
    /// <param name="currencyName"></param>
    /// <returns></returns>
    public static Currency GetCurrency(CurrencyName currencyName)
    {
        if(currencyName== CurrencyName.Euro)
            return CurrencyBuilder.CreateEuro(string.Empty);

        if (currencyName == CurrencyName.UsDollar)
            return CurrencyBuilder.CreateUsDollar(string.Empty);

        if (currencyName == CurrencyName.JapaneseYen)
            return CurrencyBuilder.CreateJapaneseYen(string.Empty);

        if (currencyName == CurrencyName.SouthKoreanWon)
            return CurrencyBuilder.CreateSouthKoreanWon(string.Empty);


        // [$CHF-417] swiss franc
        if (currencyName == CurrencyName.SwissFranc)
            return CurrencyBuilder.CreateSwissFranc(string.Empty);

        // [$¥-804]#,##0.00: chinese yuan
        if (currencyName == CurrencyName.ChineseYuan)
            return CurrencyBuilder.CreateChineseYuan(string.Empty);

        // [$£-809]#,##0.00 british pound
        if (currencyName == CurrencyName.BritishPound)        
            return CurrencyBuilder.CreateBritishPound(string.Empty);


        // [$$-C09]#,##0.00: australian dollar
        if (currencyName == CurrencyName.AustralianDollar)
            return CurrencyBuilder.CreateAustralianDollar(string.Empty);


        // [$$-1009]#,##0.00 Canadian dollar
        if (currencyName == CurrencyName.CanadianDollar)
            return CurrencyBuilder.CreateCanadianDollar(string.Empty);


        //[$$-409]#,##0.00 new zeland dollar
        if (currencyName == CurrencyName.NewZealandDollar)
            return CurrencyBuilder.CreateNewZealandDollar(string.Empty);


        //[$$-4809]#,##0.00 singapore dollar
        if (currencyName == CurrencyName.SingaporeDollar)
            return CurrencyBuilder.CreateSingaporeDollar(string.Empty);

        // [$₿]: bitcoin
        if (currencyName == CurrencyName.Bitcoin)
            return CurrencyBuilder.CreateBitcoin(string.Empty);

        // currency not manaed
        return CurrencyBuilder.CreateNotDefined(string.Empty);
    }
}
