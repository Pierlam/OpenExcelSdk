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
    /// <summary>
    /// No theme, only ARGB.
    /// </summary>
    /// <param name="argb"></param>
    public ExcelColor(string argb)
    {
        ARgb = argb;
        Rgb = "#" + GetRgb(argb);
    }

    /// <summary>
    /// Initializes a new instance of the ExcelColor class with the specified theme index, RGB color value, and tint
    /// adjustment.
    /// </summary>
    /// <param name="themeIndex">The zero-based index of the color theme to associate with this color. Must be non-negative.</param>
    /// <param name="rgb">The RGB color value as a hexadecimal string (for example, "FF0000" for red). Cannot be null or empty.</param>
    /// <param name="tint">The tint adjustment to apply to the color. Typically a value between -1.0 (darkest) and 1.0 (lightest).</param>
    public ExcelColor(int themeIndex, string rgb, double tint)
    {
        ThemeIndex = themeIndex;
        Rgb= rgb;
        Tint= tint;
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
    /// If tint is supplied, then it is applied to the RGB value of the color to determine the final color applied.
    /// The tint value is stored as a double from −1.0 .. 1.0, where −1.0 means 100% darken and 1.0 means 100% lighten.Also, 0.0 means no change.
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.colortype.tint?view=openxml-3.0.1
    /// </summary>
    public double Tint { get; set; } = 0.0;

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
