namespace OpenExcelSdk;

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