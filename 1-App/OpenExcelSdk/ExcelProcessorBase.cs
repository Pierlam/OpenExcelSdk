using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcelSdk;

/// <summary>
/// low-level ExcelProcessor functions 
/// </summary>
public class ExcelProcessorBase
{
    protected StyleMgr _styleMgr = new StyleMgr();

    #region Get CellValue

    /// <summary>
    /// Get the value of the cell as a string.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="cell"></param>
    /// <param name="stringValue"></param>
    /// <param name="error"></param>
    /// <returns></returns>
    public string GetCellValueAsString(ExcelSheet excelSheet, ExcelCell excelCell)
    {
        bool res = GetCellTypeAndValue(excelSheet, excelCell, out ExcelCellValueMulti excelCellValueMulti, out ExcelError excelError);
        if (!res)
            return string.Empty;

        if (excelCellValueMulti.CellType == ExcelCellType.String)
            return excelCellValueMulti.StringValue;

        if (excelCellValueMulti.CellType == ExcelCellType.Integer)
            return excelCellValueMulti.IntegerValue.ToString();

        if (excelCellValueMulti.CellType == ExcelCellType.Double)
            return excelCellValueMulti.DoubleValue.ToString();

        if (excelCellValueMulti.CellType == ExcelCellType.DateTime)
            return excelCellValueMulti.DateTimeValue.ToString();

        if (excelCellValueMulti.CellType == ExcelCellType.DateOnly)
            return excelCellValueMulti.DateOnlyValue.ToString();

        if (excelCellValueMulti.CellType == ExcelCellType.TimeOnly)
            return excelCellValueMulti.TimeOnlyValue.ToString();
        return string.Empty;
    }

    /// <summary>
    /// Get the value of the cell as a double.
    /// The type of the cell should match!
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="excelCell"></param>
    /// <returns></returns>
    public double GetCellValueAsDouble(ExcelSheet excelSheet, ExcelCell excelCell)
    {
        bool res = GetCellTypeAndValue(excelSheet, excelCell, out ExcelCellValueMulti excelCellValueMulti, out ExcelError excelError);
        if (!res)
            return 0;

        if (excelCellValueMulti.CellType == ExcelCellType.String)
            return 0;

        if (excelCellValueMulti.CellType == ExcelCellType.Integer)
            return excelCellValueMulti.IntegerValue.Value;

        if (excelCellValueMulti.CellType == ExcelCellType.Double)
            return excelCellValueMulti.DoubleValue.Value;

        if (excelCellValueMulti.CellType == ExcelCellType.DateTime)
            return excelCellValueMulti.DateTimeValue.Value.ToOADate();

        if (excelCellValueMulti.CellType == ExcelCellType.DateOnly)
        {
            DateTime dt = excelCellValueMulti.DateOnlyValue.Value.ToDateTime(TimeOnly.MinValue);
            return dt.ToOADate();
        }

        if (excelCellValueMulti.CellType == ExcelCellType.TimeOnly)
        {
            TimeSpan ts = excelCellValueMulti.TimeOnlyValue.Value.ToTimeSpan();
            return ts.TotalMicroseconds;
        }

        return 0;
    }

    /// <summary>
    /// Get the cell value as a date.
    /// The type of the cell should match!
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="excelCell"></param>
    /// <returns></returns>
    public DateOnly GetCellValueAsDateOnly(ExcelSheet excelSheet, ExcelCell excelCell)
    {
        bool res = GetCellTypeAndValue(excelSheet, excelCell, out ExcelCellValueMulti excelCellValueMulti, out ExcelError excelError);
        if (!res)
            return new DateOnly();

        if (excelCellValueMulti.CellType == ExcelCellType.DateOnly)
            return excelCellValueMulti.DateOnlyValue.Value;

        if (excelCellValueMulti.CellType == ExcelCellType.DateTime)
        {
            // convert the date time to date only
            return DateOnly.FromDateTime(excelCellValueMulti.DateTimeValue.Value);
        }

        return new DateOnly();
    }

    #endregion Get CellValue

    #region Get CellType

    /// <summary>
    /// Get the type of the cell value.
    /// If the cell is empty/blank, in some cases the type will be Undefined.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="excelCell"></param>
    /// <returns></returns>
    public ExcelCellType GetCellType(ExcelSheet excelSheet, ExcelCell excelCell)
    {
        bool res = GetCellTypeAndValue(excelSheet, excelCell, out ExcelCellValueMulti excelCellValueMulti, out ExcelError excelError);
        if (res) return excelCellValueMulti.CellType;
        return ExcelCellType.Error;
    }

    /// <summary>
    /// Geth the type, the value and the data format of cell.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="excelCell"></param>
    /// <param name="excelCellValueMulti"></param>
    /// <param name="excelError"></param>
    /// <returns></returns>
    public bool GetCellTypeAndValue(ExcelSheet excelSheet, ExcelCell excelCell, out ExcelCellValueMulti excelCellValueMulti, out ExcelError excelError)
    {
        excelCellValueMulti = null;
        excelError = null;

        if (excelSheet == null || excelCell == null)
        {
            excelError = new ExcelError(ExcelErrorCode.ObjectNull);
            return false;
        }

        // no cell, is null, not an error
        if (excelCell.Cell == null)
        {
            excelCellValueMulti = new ExcelCellValueMulti();
            return true;
        }

        bool isTheCase;

        //--cell datatype is defined?
        if (!GetCellStringValue(excelSheet, excelCell, out isTheCase, out excelCellValueMulti, out excelError))
            return false;
        if (isTheCase) return true;

        // get the number format id
        if (!_styleMgr.GetCellNumberFormatId(excelSheet, excelCell, out uint numberFormatId))
        {
            excelError = new ExcelError(ExcelErrorCode.TypeWrong);
            return false;
        }

        string value = excelCell.Cell.InnerText;
        double valDouble;
        int valInt;

        // is it a built-in format?
        if (BuiltInNumberFormatMgr.GetFormatAndType(numberFormatId, out string numberFormat, out ExcelCellType cellType))
        {
            return ValueBuilder.CreateValue(excelCell, cellType, value, (int)numberFormatId, numberFormat, out excelCellValueMulti, out excelError);
        }

        // Try to get custom format if exists
        if (_styleMgr.GetCustomNumberFormat(excelSheet, numberFormatId, out numberFormat))
        {
            // then determine the type from the data format: date, number,...
            cellType = GetCellTypeFromNumberFormat(numberFormat);

            return ValueBuilder.CreateValue(excelCell, cellType, value, (int)numberFormatId, numberFormat, out excelCellValueMulti, out excelError);
        }

        if (value == string.Empty)
        {
            excelCellValueMulti = new ExcelCellValueMulti();
            excelCellValueMulti.IsEmpty = true;
            return true;
        }

        // is it an int?
        bool res = int.TryParse(value, out valInt);
        if (res)
        {
            excelCellValueMulti = new ExcelCellValueMulti(valInt);
            excelCellValueMulti.Formula = excelCell.Cell?.CellFormula?.Text;
            isTheCase = true;
            return true;
        }

        // is it a double?  cultureInfo prb: replace . by ,
        value = value.Replace('.', ',');
        res = double.TryParse(value, out valDouble);
        if (res)
        {
            excelCellValueMulti = new ExcelCellValueMulti(valDouble);
            excelCellValueMulti.Formula = excelCell.Cell?.CellFormula?.Text;
            isTheCase = true;
            return true;
        }

        // not able to find the type
        excelError = new ExcelError(ExcelErrorCode.TypeWrong);
        return false;
    }

    /// <summary>
    /// Get the type of cell from the data format.
    /// exp:
    /// "dd/mm/yyyy\\ hh:mm:ss" , it's a DateTime.
    /// </summary>
    /// <param name="numberFormat"></param>
    /// <returns></returns>
    public ExcelCellType GetCellTypeFromNumberFormat(string numberFormat)
    {
        if (string.IsNullOrWhiteSpace(numberFormat))
            return ExcelCellType.Undefined;

        if ((numberFormat.Contains("y") || numberFormat.Contains("d")) && numberFormat.Contains("h"))
            return ExcelCellType.DateTime;

        if (numberFormat.Contains("y"))
            return ExcelCellType.DateOnly;

        if (numberFormat.Contains("h") || numberFormat.Contains("m"))
            return ExcelCellType.TimeOnly;

        if (numberFormat.Contains("0") || numberFormat.Contains("#"))
            return ExcelCellType.Double;

        return ExcelCellType.Undefined;
    }


    public static bool GetCellStringValue(ExcelSheet excelSheet, ExcelCell excelCell, out bool isTheCase, out ExcelCellValueMulti excelCellValueMulti, out ExcelError error)
    {
        isTheCase = false;
        excelCellValueMulti = null;
        error = null;
        string cellValue;


        if (excelCell.Cell.DataType == null) return true;

        if (excelCell.Cell.DataType.Value == CellValues.SharedString)
        {
            // SharedStringMgr
            if (!SharedStringMgr.GetSharedStringValue(excelSheet, excelCell, out cellValue))
            {
                error = new ExcelError(ExcelErrorCode.UnableGetCellStringValue);
                return false;
            }
            excelCellValueMulti = new ExcelCellValueMulti(cellValue);
            excelCellValueMulti.Formula = excelCell.Cell.CellFormula?.Text;
            isTheCase = true;
            return true;
        }

        if (excelCell.Cell.DataType.Value == CellValues.InlineString)
        {
            cellValue = excelCell.Cell.InlineString?.Text?.Text ?? string.Empty;
            if (cellValue == null)
            {
                error = new ExcelError(ExcelErrorCode.UnableGetCellStringValue);
                return false;
            }
            excelCellValueMulti = new ExcelCellValueMulti(cellValue);
            excelCellValueMulti.Formula = excelCell.Cell.CellFormula?.Text;
            isTheCase = true;
            return true;
        }

        if (excelCell.Cell.DataType.Value == CellValues.String)
        {
            string value = excelCell.Cell.InnerText;
            if (value == null) value = string.Empty;
            excelCellValueMulti = new ExcelCellValueMulti(value);
            excelCellValueMulti.Formula = excelCell.Cell.CellFormula?.Text;
            isTheCase = true;
            return true;
        }

        // not a string, bye
        return true;
    }

    #endregion Get Cell, CellValue, CellType

    #region Create cell

    /// <summary>
    /// Given a column name, a row index, and a WorksheetPart, inserts a cell into the worksheet.
    /// If the cell already exists, returns it.
    /// Exp: "A", 1 for cell A1.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="columnName"></param>
    /// <param name="rowIndex"></param>
    /// <param name="excelCell"></param>
    /// <param name="error"></param>
    /// <returns></returns>
    public bool CreateCell(ExcelSheet excelSheet, string columnName, uint rowIndex, out ExcelCell excelCell, out ExcelError error)
    {
        excelCell = null;
        error = null;

        Worksheet worksheet = excelSheet.Worksheet;
        SheetData? sheetData = worksheet.GetFirstChild<SheetData>();
        string cellReference = columnName + rowIndex;

        // If the worksheet does not contain a row with the specified row index, insert one.
        Row row;

        if (sheetData?.Elements<Row>().Where(r => r.RowIndex is not null && r.RowIndex == rowIndex).Count() != 0)
        {
            row = sheetData!.Elements<Row>().Where(r => r.RowIndex is not null && r.RowIndex == rowIndex).First();
        }
        else
        {
            row = new Row() { RowIndex = rowIndex };
            sheetData.Append(row);
        }

        // If there is not a cell with the specified column name, insert one.
        if (row.Elements<Cell>().Where(c => c.CellReference is not null && c.CellReference.Value == columnName + rowIndex).Count() > 0)
        {
            Cell cell = row.Elements<Cell>().Where(c => c.CellReference is not null && c.CellReference.Value == cellReference).First();
            excelCell = new ExcelCell(excelSheet, cell);
            return true;
        }
        else
        {
            // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
            Cell? refCell = null;

            foreach (Cell cell in row.Elements<Cell>())
            {
                if (string.Compare(cell.CellReference?.Value, cellReference, true) > 0)
                {
                    refCell = cell;
                    break;
                }
            }

            Cell newCell = new Cell() { CellReference = cellReference };
            row.InsertBefore(newCell, refCell);

            excelCell = new ExcelCell(excelSheet, newCell);
            return true;
        }
    }

    #endregion Create cell

    #region Set cell value

    public bool SetCellValue(ExcelSheet excelSheet, ExcelCell excelCell, string value, string format, out ExcelError error)
    {
        error = null;
        uint formatId;

        // get (built-in or custom) or create the format (custom)
        if (!_styleMgr.GetOrCreateNumberFormat(excelSheet, format, out formatId, out error))
            return false;

        return SetCellValueAndNumberFormatId(excelSheet, excelCell, value, formatId, out error);
    }

    public bool SetCellValue(ExcelSheet excelSheet, ExcelCell excelCell, int value, string format, out ExcelError error)
    {
        error = null;
        uint formatId;

        // get (built-in or custom) or create the format (custom)
        if (!_styleMgr.GetOrCreateNumberFormat(excelSheet, format, out formatId, out error))
            return false;

        return SetCellValueAndNumberFormatId(excelSheet, excelCell, value, formatId, out error);
    }

    public bool SetCellValue(ExcelSheet excelSheet, ExcelCell excelCell, double value, string format, out ExcelError error)
    {
        error = null;
        uint formatId;

        // get (built-in or custom) or create the format (custom)
        if (!_styleMgr.GetOrCreateNumberFormat(excelSheet, format, out formatId, out error))
            return false;

        return SetCellValueAndNumberFormatId(excelSheet, excelCell, value, formatId, out error);
    }

    public bool SetCellValue(ExcelSheet excelSheet, ExcelCell excelCell, DateOnly value, string format, out ExcelError error)
    {
        error = null;
        uint formatId;

        // get (built-in or custom) or create the format (custom)
        if (!_styleMgr.GetOrCreateNumberFormat(excelSheet, format, out formatId, out error))
            return false;

        return SetCellValueAndNumberFormatId(excelSheet, excelCell, value, formatId, out error);
    }

    public bool SetCellValue(ExcelSheet excelSheet, ExcelCell excelCell, DateTime value, string format, out ExcelError error)
    {
        error = null;
        uint formatId;

        // get (built-in or custom) or create the format (custom)
        if (!_styleMgr.GetOrCreateNumberFormat(excelSheet, format, out formatId, out error))
            return false;

        return SetCellValueAndNumberFormatId(excelSheet, excelCell, value, formatId, out error);
    }

    public bool SetCellValue(ExcelSheet excelSheet, ExcelCell excelCell, TimeOnly value, string format, out ExcelError error)
    {
        error = null;
        uint formatId;

        // get (built-in or custom) or create the format (custom)
        if (!_styleMgr.GetOrCreateNumberFormat(excelSheet, format, out formatId, out error))
            return false;

        return SetCellValueAndNumberFormatId(excelSheet, excelCell, value, formatId, out error);
    }

    public bool SetCellValue(ExcelSheet excelSheet, ExcelCell excelCell, string value, out ExcelError error)
    {
        return SetCellValueAndNumberFormatId(excelSheet, excelCell, value, null, out error);
    }

    public bool SetCellValue(ExcelSheet excelSheet, ExcelCell excelCell, int value, out ExcelError error)
    {
        return SetCellValueAndNumberFormatId(excelSheet, excelCell, value, null, out error);
    }

    public bool SetCellValue(ExcelSheet excelSheet, ExcelCell excelCell, double value, out ExcelError error)
    {
        return SetCellValueAndNumberFormatId(excelSheet, excelCell, value, null, out error);
    }

    #endregion Set cell value

    #region Set cell value and number format Id

    /// <summary>
    /// Set a string value and a number format in the existing cell.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="excelCell"></param>
    /// <param name="value"></param>
    /// <param name="error"></param>
    /// <returns></returns>
    public bool SetCellValueAndNumberFormatId(ExcelSheet excelSheet, ExcelCell excelCell, string value, uint? numberFormatId, out ExcelError error)
    {
        error = null;

        if (excelCell == null || excelCell.Cell == null)
        {
            error = new ExcelError(ExcelErrorCode.ObjectNull);
            return false;
        }

        try
        {
            WorkbookPart workbookPart = excelSheet.ExcelFile.WorkbookPart;

            // get the table
            SharedStringTablePart shareStringPart = SharedStringMgr.GetOrCreateSharedStringTablePart(excelSheet.ExcelFile.WorkbookPart);

            // Insert the text into the SharedStringTablePart
            int index = SharedStringMgr.InsertSharedStringItem(value, shareStringPart);

            // Set the value of cell A1
            excelCell.Cell.CellValue = new CellValue(index.ToString());
            excelCell.Cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);

            // remove formula if it's there
            _styleMgr.RemoveFormula(excelSheet, excelCell);

            // no cell format, nothing more to do
            if (numberFormatId == null && !_styleMgr.HasCellFormat(excelSheet, excelCell)) return true;

            // the cell contains the expected number format
            _styleMgr.GetCellNumberFormatId(excelSheet, excelCell, out uint numberFormatIdCell);
            if (numberFormatIdCell == (numberFormatId ?? 0)) return true;

            // all other style than format (no border, no color,...) are null, clear the style of the cell
            if (numberFormatId == null && _styleMgr.AllOthersStyleThanFormatAreNull(excelSheet, excelCell))
            {
                // no format to set, all others style part style are null, so clear the style
                excelCell.Cell.StyleIndex = 0;
                return true;
            }

            // duplicate the style to update the CellFormat
            _styleMgr.UpdateCellStyleNumberFormatId(excelSheet, excelCell, numberFormatId);

            return true;
        }
        catch (Exception ex)
        {
            error = new ExcelError(ExcelErrorCode.UnableSetCellValue, ex);
            return false;
        }
    }

    /// <summary>
    /// Set cell value as double.
    /// Keep some aprt of the style: border, color, font...
    /// but clear the number format -> style/CellFormat/NumberingFormat
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="excelCell"></param>
    /// <param name="value"></param>
    /// <param name="error"></param>
    /// <returns></returns>
    public bool SetCellValueAndNumberFormatId(ExcelSheet excelSheet, ExcelCell excelCell, double value, uint? numberFormatId, out ExcelError error)
    {
        error = null;

        // Important: store as number
        excelCell.Cell.DataType = CellValues.Number;
        excelCell.Cell.CellValue = new CellValue(value.ToString(global::System.Globalization.CultureInfo.InvariantCulture));

        // remove formula if it's there
        _styleMgr.RemoveFormula(excelSheet, excelCell);

        // no cell format, nothing more to do
        if (numberFormatId == null && !_styleMgr.HasCellFormat(excelSheet, excelCell)) return true;

        // the cell contains the expected number format
        _styleMgr.GetCellNumberFormatId(excelSheet, excelCell, out uint numberFormatIdCell);
        if (numberFormatIdCell == (numberFormatId ?? 0)) return true;

        // all other style than format (no border, no color,...) are null, clear the style of the cell
        if (numberFormatId == null && _styleMgr.AllOthersStyleThanFormatAreNull(excelSheet, excelCell))
        {
            // no format to set, all others style part style are null, so clear the style
            excelCell.Cell.StyleIndex = 0;
            return true;
        }

        // duplicate the style to update the CellFormat
        _styleMgr.UpdateCellStyleNumberFormatId(excelSheet, excelCell, numberFormatId);

        return true;
    }

    /// <summary>
    /// Set cell value as a date.
    /// Keep some aprt of the style: border, color, font...
    /// but clear the number format -> style/CellFormat/NumberingFormat
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="excelCell"></param>
    /// <param name="value"></param>
    /// <param name="error"></param>
    /// <returns></returns>
    public bool SetCellValueAndNumberFormatId(ExcelSheet excelSheet, ExcelCell excelCell, DateOnly value, uint? numberFormatId, out ExcelError error)
    {
        error = null;

        // Important: store as number
        excelCell.Cell.DataType = CellValues.Number;

        // Convert to DateTime (midnight time)
        DateTime dateTime = value.ToDateTime(TimeOnly.MinValue);

        // Convert to double (OLE Automation date)
        double oaDate = dateTime.ToOADate();
        excelCell.Cell.CellValue = new CellValue(oaDate.ToString(global::System.Globalization.CultureInfo.InvariantCulture));

        // remove formula if it's there
        _styleMgr.RemoveFormula(excelSheet, excelCell);

        // numberFormatId is mandatory for date
        if (numberFormatId == null)
        {
            error = new ExcelError(ExcelErrorCode.FormatMissingForDate);
            return false;
        }

        // the cell contains the expected number format
        _styleMgr.GetCellNumberFormatId(excelSheet, excelCell, out uint numberFormatIdCell);
        if (numberFormatIdCell == (numberFormatId ?? 0)) return true;

        // duplicate the style to update the CellFormat
        _styleMgr.UpdateCellStyleNumberFormatId(excelSheet, excelCell, numberFormatId);

        return true;
    }

    /// <summary>
    /// Set cell value as a date.
    /// Keep some aprt of the style: border, color, font...
    /// but clear the number format -> style/CellFormat/NumberingFormat
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="excelCell"></param>
    /// <param name="value"></param>
    /// <param name="error"></param>
    /// <returns></returns>
    public bool SetCellValueAndNumberFormatId(ExcelSheet excelSheet, ExcelCell excelCell, DateTime value, uint? numberFormatId, out ExcelError error)
    {
        error = null;

        // Important: store as number
        excelCell.Cell.DataType = CellValues.Number;

        // Convert to double (OLE Automation date)
        double oaDate = value.ToOADate();
        excelCell.Cell.CellValue = new CellValue(oaDate.ToString(global::System.Globalization.CultureInfo.InvariantCulture));

        // remove formula if it's there
        _styleMgr.RemoveFormula(excelSheet, excelCell);

        // numberFormatId is mandatory for date
        if (numberFormatId == null)
        {
            error = new ExcelError(ExcelErrorCode.FormatMissingForDate);
            return false;
        }

        // the cell contains the expected number format
        _styleMgr.GetCellNumberFormatId(excelSheet, excelCell, out uint numberFormatIdCell);
        if (numberFormatIdCell == (numberFormatId ?? 0)) return true;

        // duplicate the style to update the CellFormat
        _styleMgr.UpdateCellStyleNumberFormatId(excelSheet, excelCell, numberFormatId);

        return true;
    }

    public bool SetCellValueAndNumberFormatId(ExcelSheet excelSheet, ExcelCell excelCell, TimeOnly value, uint? numberFormatId, out ExcelError error)
    {
        error = null;

        // Important: store as number
        excelCell.Cell.DataType = CellValues.Number;

        // set the hour, minute, second and millisecond
        DateTime dateTime = new DateTime(2025, 1, 1, value.Hour, value.Minute, value.Second, value.Millisecond);

        double oaDate = dateTime.ToOADate();
        // get the fractional part only
        oaDate = oaDate - Math.Truncate(oaDate);

        // Convert to double (OLE Automation date)
        excelCell.Cell.CellValue = new CellValue(oaDate.ToString(global::System.Globalization.CultureInfo.InvariantCulture));

        // remove formula if it's there
        _styleMgr.RemoveFormula(excelSheet, excelCell);

        // numberFormatId is mandatory for date
        if (numberFormatId == null)
        {
            error = new ExcelError(ExcelErrorCode.FormatMissingForDate);
            return false;
        }

        // the cell contains the expected number format
        _styleMgr.GetCellNumberFormatId(excelSheet, excelCell, out uint numberFormatIdCell);
        if (numberFormatIdCell == (numberFormatId ?? 0)) return true;

        // duplicate the style to update the CellFormat
        _styleMgr.UpdateCellStyleNumberFormatId(excelSheet, excelCell, numberFormatId);

        return true;
    }

    #endregion Set cell value and number format Id

}
