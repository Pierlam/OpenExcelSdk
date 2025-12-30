using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcelSdk;

/// <summary>
/// Excel exception.
/// </summary>
public class ExcelException :Exception
{
    /// <summary>
    /// the method name where the error occurs.
    /// </summary>
    public string Action {  get; set; }

    public ExcelErrorCode ExcelErrorCode { get; set;  }

    public ExcelException(string action, ExcelErrorCode excelErrorCode)
    {
        Action = action;
        ExcelErrorCode = excelErrorCode;
    }

    public ExcelException(string action, ExcelErrorCode excelErrorCode, string message)
        : base(message)
    {
        Action = action;
        ExcelErrorCode = excelErrorCode;
    }

    public ExcelException(string action, ExcelErrorCode excelErrorCode, string message, Exception inner)
        : base(message, inner)
    {
        Action = action;
        ExcelErrorCode = excelErrorCode;
    }

    public static ExcelException Create(string action, ExcelErrorCode excelErrorCode)
    {
        return Create(action, excelErrorCode, string.Empty);
    }

    public static ExcelException Create(string action, ExcelErrorCode excelErrorCode, string param)
    {
        string msg= ErrorMsgBuilder.BuildMsg(action, excelErrorCode, param);
        return new ExcelException(action, excelErrorCode, msg);
    }

    public static ExcelException Create(string action, ExcelErrorCode excelErrorCode, string param, Exception ex)
    {
        string msg = ErrorMsgBuilder.BuildMsg(action, excelErrorCode, param);
        return new ExcelException(action, excelErrorCode, msg, ex);
    }

}
