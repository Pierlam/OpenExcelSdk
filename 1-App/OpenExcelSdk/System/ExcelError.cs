using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcelSdk.System;
public enum ExcelErrorCode
{
    NoError,

    ObjectNull,
    ValueNull,
    FileNull,

    UnableCreateFile,
    FileAlreadyExists,
    FileNotFound,
    UnableOpenFile,
    UnableCloseFile,
    
    UnableGetSheet,

    IndexWrong,
    IndexMustBePositive,
    TypeWrong,

    UnableGetCell,
    UnableGetRow,

    UnableGetCellStringValue,    
    UnableSetCellValue,
    FormatMissingForDate,

    UnableRemoveCell
}

public class ExcelError
{
    public ExcelError(ExcelErrorCode errorCode)
    {
        ErrorCode = errorCode;
    }
    public ExcelError(ExcelErrorCode errorCode, Exception exception)
    {
        ErrorCode = errorCode;
        Exception = exception;
    }

    public ExcelErrorCode ErrorCode { get; set; } = ExcelErrorCode.NoError;

    public string Message { get; set; } = string.Empty;
    public Exception? Exception { get; set; } = null;
}
