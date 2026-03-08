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
    /// Represents a currency, including its symbol. 
    /// Negative sign is displayed.
    /// e.g.: #,##0.00\ "€"
    /// </summary>
    Currency,

    /// <summary>
    /// Represents a currency, including its symbol.
    /// Negative numbers are displayed in red color. Negative sign is displayed.
    /// e.g.: #,##0.00\ "€";[Red]\-#,##0.00\ "€"
    /// </summary>
    CurrencyRedNegative,

    /// <summary>
    /// Represents a currency, including its symbol.
    /// Negative numbers are displayed in red color. Negative sign is not displayed.
    /// e.g.: #,##0.00\ "€";[Red]#,##0.00\ "€"
    /// </summary>
    CurrencyRedNegativeNoSign,

    /// <summary>
    /// Represents a currency, including its symbol. 
    /// Negative sign is displayed.
    /// e.g.:  xxx 
    /// </summary>
    CurrencyLeftSpace,

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
    public string ExcelCode { get; set; }=string.Empty;

    /// <summary>
    /// Display format: Accounting, Currency, CurrencyRedNegative, CurrencyRedNegativeNoSign, CurrencyLeftSpace.
    /// </summary>
    public CurrencyFormat Format {  get; set; }= CurrencyFormat.Currency;

    /// <summary>
    /// Gets or sets the position of the currency symbol relative to the numeric amount.
    /// e.g. for Euro, the symbol is typically placed after the amount (e.g., 100 €), while for US Dollar, the symbol is typically placed before the amount (e.g., $100).
    /// </summary>
    /// <remarks>The default value is <see cref="CurrencySymbolPosition.Before"/>, which places the currency
    /// symbol before the amount. This property can be adjusted to match regional or formatting preferences.</remarks>
    public CurrencySymbolPosition SymbolPosition { get;set;  }= CurrencySymbolPosition.Before;
}
