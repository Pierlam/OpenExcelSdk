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

    Bitcoin
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
}

