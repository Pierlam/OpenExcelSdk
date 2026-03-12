namespace OpenExcelSdk;

/// <summary>
/// To build error message based on a code, used in exception.
/// </summary>
public class ErrorMsgBuilder
{
    public static string BuildMsg(string action, ExcelErrorCode errorCode, string param)
    {
        if (string.IsNullOrWhiteSpace(action)) action = "(fct)";
        if (string.IsNullOrWhiteSpace(param)) param = string.Empty;

        if (errorCode == ExcelErrorCode.FilenameNull) return String.Format("FilenameNull: {0}: The filename is null.", action);
        if (errorCode == ExcelErrorCode.FileAlreadyExists) return String.Format("FileAlreadyExists: {0}: The filename {1} already exists.", action, param);
        if (errorCode == ExcelErrorCode.FileNotFound) return String.Format("FileNotFound: {0}: The file {1} is not found.", action, param);
        if (errorCode == ExcelErrorCode.UnableOpenFile) return String.Format("UnableOpenFile: {0}: Unble to open the file {1} is not found.", action, param);
        if (errorCode == ExcelErrorCode.UnableCreateFile) return String.Format("UnableCreateFile: {0}: Unble to create a new file {1} .", action, param);
        if (errorCode == ExcelErrorCode.UnableCloseFile) return String.Format("UnableCloseFile: {0}: Unble to close the file {1}.", action, param);

        if (errorCode == ExcelErrorCode.UnableCreateSheet) return String.Format("UnableCreateSheet: {0}: Unable to create a new sheet in the Excel file.", action);
        if (errorCode == ExcelErrorCode.UnableGetSheet) return String.Format("UnableGetSheet: {0}: Unable to get a sheet from the Excel file, sheet index/name: {1}", action, param);
        if (errorCode == ExcelErrorCode.SheetnameNull) return String.Format("SheetnameNull: {0}: The sheet name is null.", action);

        if (errorCode == ExcelErrorCode.UnableGetRow) return String.Format("UnableGetRow: {0}: Unable to get a row from the Excel sheet, index: {1}", action, param);

        if (errorCode == ExcelErrorCode.UnableGetCell) return String.Format("UnableGetCell: {0}: Unable to get a cell from the Excel sheet, cell address: {1}", action, param);
        if (errorCode == ExcelErrorCode.UnableGetCellStringValue) return String.Format("UnableGetCellStringValue: {0}: Unable to get the shared string from the table.", action);

        if (errorCode == ExcelErrorCode.ObjectNull) return String.Format("ObjectNull: {0}: The object is null, can be: sheet, cell, row...", action);
        if (errorCode == ExcelErrorCode.TypeWrong) return String.Format("TypeWrong: {0}: The type of the value is wrong or not expected, value: {1}", action, param);

        if (errorCode == ExcelErrorCode.IndexWrong) return String.Format("IndexWrong: {0}: The index is wrong/not expected, value: {1}", action, param);
        if (errorCode == ExcelErrorCode.IndexMustBePositive) return String.Format("IndexMustBePositive: {0}: The index must be positive, value: {1}", action, param);
        if (errorCode == ExcelErrorCode.UnableRemoveCell) return String.Format("UnableRemoveCell: {0}: Unable to remove the cell, address: {1}", action, param);

        // if the code is not managed
        return String.Format("{0}: An internal error occurs, code: {1}", action, errorCode);
    }
}