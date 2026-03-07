using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcelSdk.System;

public enum CurrencyCode
{
    // not defined
    XXX,

    EUR,
    USD,
    GBP,
    CHF,
    JPY,
    KWR,
    AUD,
    CAD,
    CNY,
    NZD,
    SGD,
    BTC
}

public enum CurrencyName
{
    NotDefined,
    Euro,
    UsDollar,
    BritishPound,
    SwissFranc,
    JapaneseYen,
    SouthKoreanWon,
    AustralianDollar,
    CanadianDollar,
    ChineseYuan,
    NewZealandDollar,
    SingaporeDollar,
    Bitcoin
}

public enum CurrencySymbolPosition
{

    /// <summary>
    /// The symbol is placed before the numeric value, which is a common convention for many currencies.
    /// e.g. for many currency like dollar.
    /// </summary>
    Before,

    /// <summary>
    /// The symbol is placed after the numeric value, which is a common convention for many currencies.
    /// e.g. for Euro.
    /// </summary>
    /// <remarks>This property is typically used to assess the outcome of an operation and may influence
    /// subsequent actions or decisions based on its value.</remarks>
    After
}

public enum CurrencyFormat
{
    /// <summary>
    /// Represents a currency, including its code, symbol, and formatting conventions for monetary values.
    /// e.g. NumberFormatId=164	#,##0.00\ "€"
    /// negative value are displayed in red color.
    /// </summary>
    Currency,

    /// <summary>
    /// Currency Symbol is displayed on left side.
    /// e.g. NumberFormatId=44	_-* #,##0.00\ "€"_-;\-* #,##0.00\ "€"_-;_-* "-"??\ "€"_-;_-@_-
    /// </summary>
    Accounting
}


/// <summary>
/// Currency definition.
/// </summary>
public class Currency
{
    /// <summary>
    /// Currency symbol, e.g. $, €, ¥
    /// </summary>
    public string Symbol { get; set; }

    /// <summary>
    /// Currency code, e.g. USD, EUR
    /// </summary>
    public CurrencyCode Code { get; set; }

    /// <summary>
    /// Currency name, e.g. UsDollar, Euro
    /// </summary>
    public CurrencyName Name { get; set; }

    /// <summary>
    /// code used in excel format, e.g. for Euro: "€"
    /// [$$-409]  for US Dollar, ...
    /// </summary>
    public string ExcelCode { get; set; }

    public CurrencySymbolPosition SymbolPosition { get;set;  }= CurrencySymbolPosition.Before;
}

