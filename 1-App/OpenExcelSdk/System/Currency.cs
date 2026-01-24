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
    public string Symbol { get; set; } 
    public CurrencyCode Code { get; set; }

    public CurrencyName Name { get; set; }
}

