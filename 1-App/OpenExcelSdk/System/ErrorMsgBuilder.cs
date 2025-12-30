using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcelSdk;

public class ErrorMsgBuilder
{
    public static string BuildMsg(string action, ExcelErrorCode errorCode, string param)
    {
        if (string.IsNullOrWhiteSpace(param)) param = string.Empty;

        if (errorCode == ExcelErrorCode.FilenameNull) return String.Format("{0}: The filename is null.", action); 
        if (errorCode == ExcelErrorCode.FileAlreadyExists) return String.Format("{0}: The filename {0} already exists.", action, param);

        return String.Format("{0}: An internal error occurs, code: ", errorCode);
    }

}
