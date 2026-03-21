using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcelExport;

public class HelpPrinter
{
    public static void PrintHelp(string vers)
    {
        Console.WriteLine();
        Console.WriteLine("=> OpenExcelExport - Excel styles exporter");
        Console.WriteLine();
        Console.WriteLine("Version: " + vers + " by Pierlam, March 2026");
        Console.WriteLine();
        Console.WriteLine("Website: https://pierlam.github.io/OpenExcelExport/");
        Console.WriteLine();
        Console.WriteLine("Goal:");
        Console.WriteLine("  Axport all styles from an input excel file: Shared strings, CellFormat, NumberFormat, Fill, Font,...");
        Console.WriteLine("  Result is saved into an output excel file.");
        Console.WriteLine();
        Console.WriteLine("Usage:");
        Console.WriteLine("  OpenExcelExport -excel <input_excel_file> -out <output_excel_file>");
        Console.WriteLine();
        Console.WriteLine("Parameters:");
        Console.WriteLine("  -excel : Full path of the input excel file to analyze.");
        Console.WriteLine("  -out   : Full path of the output excel file to create.");
        Console.WriteLine();
        Console.WriteLine("Remark:");
        Console.WriteLine("  If filename contains space, add \" around it.");
        Console.WriteLine();
        Console.WriteLine("Examples:");
        Console.WriteLine("  OpenExcelExport -excel C:\\Input\\source.xlsx -out result.xlsx");
        Console.WriteLine("  OpenExcelExport -excel \"C:\\My Files\\source.xlsx\" -out result.xlsx");
        Console.WriteLine();
    }
}
