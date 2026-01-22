using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcelExport;

public class ArgsParser
{
    /// <summary>
    /// Parse the arguments line.
    /// </summary>
    /// <param name="arg"></param>
    /// <returns></returns>
    public static bool Parse(string arg, out ProgParams progParams, out string errMsg)
    {
        progParams= new ProgParams();
        arg = arg.Trim();

        string argEnd;

        if (!CheckRemove(arg, "-excel", out argEnd, out errMsg))
            return false;

        RemoveSpace(argEnd, out argEnd);

        if (!CheckRemove(argEnd, "=", out argEnd, out errMsg))
            return false;

        RemoveSpace(argEnd, out argEnd);

        if (!CheckRemoveString(argEnd, out argEnd, out string filenameIn, out errMsg))
            return false;

        progParams.InputExcelFile = filenameIn;

        if (!CheckRemoveSpace(argEnd, out argEnd, out errMsg))
            return false;

        if (!CheckRemove(argEnd, "-out", out argEnd, out errMsg))
            return false;

        RemoveSpace(argEnd, out argEnd);

        if (!CheckRemove(argEnd, "=", out argEnd, out errMsg))
            return false;

        RemoveSpace(argEnd, out argEnd);

        if (!CheckRemoveString(argEnd, out argEnd, out string filenameOut, out errMsg))
            return false;

        progParams.OutputExcelFile= filenameOut;
        return true;
    }


    public static bool CheckRemove(string arg, string item, out string argEnd, out string errMsg)
    {
        argEnd= string.Empty;   
        errMsg = string.Empty;

        // the item should be here at the start
        if (!arg.StartsWith(item))
        {
            errMsg= "Error, argument: " + item + " expected, but not found.";
            return false;
        }

        // remove the item
        argEnd = arg.Substring(item.Length);

        //if (argEnd.Length == 0)
        //{
        //    errMsg = "Error, argument: " + item + " expected, but not found.";
        //    return false;
        //}

        return true;
    }

    public static bool CheckRemoveString(string arg, out string argEnd, out string paramString, out string errMsg)
    {
        argEnd = string.Empty;
        errMsg = string.Empty;
        paramString = string.Empty;

        // get the start quote
        if (!CheckRemove(arg, "'", out argEnd, out errMsg))
            return false;

        // get the string until the end quote
        int quoteEndPos = argEnd.IndexOf("'");
        if(quoteEndPos<0)
        {
            errMsg= "Error, missing end quote for string : " + argEnd;
            return false;
        }

        paramString= argEnd.Substring(0, quoteEndPos);
        argEnd= argEnd.Substring(quoteEndPos + 1);
        return true;
    }

    /// <summary>
    /// At least one space should be here and removed.
    /// </summary>
    /// <param name="arg"></param>
    /// <param name="argEnd"></param>
    /// <param name="errMsg"></param>
    /// <returns></returns>
    public static bool CheckRemoveSpace(string arg, out string argEnd, out string errMsg)
    {
        argEnd = string.Empty;
        errMsg = string.Empty;

        if(arg.Length==0 || char.IsWhiteSpace(arg[0])==false)
        {
            errMsg= "Error, space expected, but not found, argument : " +arg;
            return false;
        }
        argEnd= arg.Trim();
        return true;
    }

    public static void RemoveSpace(string arg, out string argEnd)
    {
        argEnd = arg.Trim();
    }

}
