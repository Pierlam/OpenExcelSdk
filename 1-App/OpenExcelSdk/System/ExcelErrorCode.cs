namespace OpenExcelSdk;


/// <summary>
/// Excel error code.
/// </summary>
public enum ExcelErrorCode
{
    //NoError,

    ObjectNull,
    //ValueNull,

    FilenameNull,
    FileNotFound,

    UnableCreateFile,
    FileAlreadyExists,
    UnableOpenFile,
    UnableCloseFile,

    SheetnameNull,
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
