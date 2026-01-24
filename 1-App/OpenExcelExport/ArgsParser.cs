using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcelExport;

public class ArgsParser
{
    public static bool Parse(string[] args, out ProgParams progParams, out string errMsg)
    {
        progParams= new ProgParams();

        if (args.Length!=4)
        {
            errMsg= "Error, 4 arguments expected.";
            progParams= null;
            return false;
        }

        // -excel
        if(args[0].ToLower()!="-excel")
        {
            errMsg= "Error, first argument should be -excel.";
            progParams= null;
            return false;
        }

        // excel filename to analyze
        if (args[1].Trim().Length == 0)
        {
            errMsg = "Error, filename to analyze is empty.";
            progParams = null;
            return false;
        }
        progParams.InputExcelFile = args[1].Trim();

        // -out
        if (args[2].ToLower() != "-out")
        {
            errMsg = "Error, Third argument should be -out.";
            progParams = null;
            return false;
        }

        // excel filename to analyze
        if (args[3].Trim().Length == 0)
        {
            errMsg = "Error, output filename is empty.";
            progParams = null;
            return false;
        }
        progParams.OutputExcelFile = args[3].Trim();
        errMsg=string.Empty;
        return true;
    }
}
