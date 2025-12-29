namespace OpenExcelSdk;


/// <summary>
/// Excel error code.
/// </summary>
public enum ExcelErrorCode
{
    NoError,

    ObjectNull,
    ValueNull,
    FilenameNull,

    FileNotFound,

    UnableCreateFile,
    FileAlreadyExists,
    UnableOpenFile,
    UnableCloseFile,

    UnableCreateSheet,
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

/// <summary>
/// Excel error object.
/// </summary>
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