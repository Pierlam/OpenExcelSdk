using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcelSdk.System;

/// <summary>
/// Background or foreground color.
/// </summary>
public class ExcelColor
{
    public ExcelColor(string argb)
    {
        ARgb = argb;
        Rgb = "#" + GetRgb(argb);
    }

    public ExcelColor(int themeIndex, string rgb)
    {
        ThemeIndex = themeIndex;
        Rgb= rgb;
    }

    /// <summary>
    /// Addressable RGB.
    /// ARGB format: FF (opaque) + RGB.
    /// exp: "FFFFFF00" which is yellow.
    /// Can be empty is the color is defined in theme.
    /// </summary>
    public string ARgb { get; set; } = string.Empty;

    /// <summary>
    /// The rgb value, with the prefix.
    /// exp; #FFFF00, yellow.
    /// Always set.
    /// </summary>
    public string Rgb { get; set; } = string.Empty;

    /// <summary>
    /// Theme index if the color is coming from the theme (predefined colors).
    /// </summary>
    public int ThemeIndex { get; set; } = -1;

    /// <summary>
    /// Set a color name if exists.
    /// </summary>
    public ColorName ColorName { get; set; }=ColorName.Undefined;

    /// <summary>
    /// Get hte rgb value, remove the prefix Addressable/Alpha.
    /// </summary>
    /// <param name="argb"></param>
    /// <returns></returns>
    public string GetRgb(string argb)
    {
        if(string.IsNullOrWhiteSpace(argb))return string.Empty;
        if (argb.Length < 2) return argb;
        return argb.Substring(2,argb.Length-2);
    }
}
