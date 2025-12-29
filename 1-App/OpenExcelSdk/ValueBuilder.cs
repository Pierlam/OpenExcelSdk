namespace OpenExcelSdk;

/// <summary>
/// To build ExcelCellValueMulti object.
/// type and value of a cell.
/// </summary>
public class ValueBuilder
{
    public static bool CreateValue(ExcelCell excelCell, ExcelCellType cellType, string value, int numberFormatId, string numberFormat, out ExcelCellValueMulti excelCellValueMulti, out ExcelError excelError)
    {
        excelCellValueMulti = null;
        excelError = null;

        if (string.IsNullOrEmpty(value))
        {
            excelCellValueMulti = new ExcelCellValueMulti();
            excelCellValueMulti.CellType = cellType;
            excelCellValueMulti.IsEmpty = true;
            return true;
        }

        if (cellType == ExcelCellType.Integer)
        {
            if(!ValueBuilder.CreateValueInteger(value, (int)numberFormatId, numberFormat, out excelCellValueMulti, out excelError))
                return false;
            excelCellValueMulti.Formula = excelCell.Cell.CellFormula?.Text;
            return true;
        }

        if (cellType == ExcelCellType.Double)
        {
            if (!ValueBuilder.CreateValueDouble(value, (int)numberFormatId, numberFormat, out excelCellValueMulti, out excelError))
                return false;
            excelCellValueMulti.Formula = excelCell.Cell.CellFormula?.Text;
            return true;
        }

        if (cellType == ExcelCellType.DateOnly)
        {
            if (!ValueBuilder.CreateValueDateOnly(value, (int)numberFormatId, numberFormat, out excelCellValueMulti, out excelError))
                return false;
            excelCellValueMulti.Formula = excelCell.Cell.CellFormula?.Text;
            return true;
        }

        if (cellType == ExcelCellType.DateTime)
        {
            if(!ValueBuilder.CreateValueDateTime(value, (int)numberFormatId, numberFormat, out excelCellValueMulti, out excelError))
                return false;
            excelCellValueMulti.Formula = excelCell.Cell.CellFormula?.Text;
            return true;
        }

        if (cellType == ExcelCellType.TimeOnly)
        {
            if(!ValueBuilder.CreateValueTimeOnly(value, (int)numberFormatId, numberFormat, out excelCellValueMulti, out excelError))
                return false;
            excelCellValueMulti.Formula = excelCell.Cell.CellFormula?.Text;
            return true;
        }

        excelError = new ExcelError(ExcelErrorCode.TypeWrong);
        return false;
    }

    public static bool CreateValueInteger(string value, int numberFormatId,  string numberFormat, out ExcelCellValueMulti excelCellValueMulti, out ExcelError excelError)
    {
        bool res = int.TryParse(value, out int valInt);
        if (!res)
        {
            excelError = new ExcelError(ExcelErrorCode.TypeWrong);
            excelCellValueMulti = null;
            return false;
        }
        excelError = null;
        excelCellValueMulti = new ExcelCellValueMulti(valInt);
        excelCellValueMulti.NumberFormatId = numberFormatId;
        excelCellValueMulti.NumberFormat = numberFormat;
        return true;
    }

    public static bool CreateValueDouble(string value, int numberFormatId, string numberFormat, out ExcelCellValueMulti excelCellValueMulti, out ExcelError excelError)
    {
        // cultureInfo prb: replace . by ,
        value = value.Replace('.', ',');
        bool res = double.TryParse(value, out double valDouble);
        if (!res)
        {
            excelError = new ExcelError(ExcelErrorCode.TypeWrong);
            excelCellValueMulti = null;
            return false;
        }
        excelError = null;
        excelCellValueMulti = new ExcelCellValueMulti(valDouble);
        excelCellValueMulti.NumberFormatId = numberFormatId;
        excelCellValueMulti.NumberFormat = numberFormat;
        return true;
    }

    public static bool CreateValueDateOnly(string value, int numberFormatId, string numberFormat, out ExcelCellValueMulti excelCellValueMulti, out ExcelError excelError)
    {
        excelError = null;

        try
        {
            value = value.Replace('.', ',');
            double valDouble = double.Parse(value);

            // convert the value to date
            DateTime dateTime = DateTime.FromOADate(valDouble);
            DateOnly dateOnly = DateOnly.FromDateTime(dateTime);
            excelCellValueMulti = new ExcelCellValueMulti(dateOnly);
            excelCellValueMulti.NumberFormatId = numberFormatId;
            excelCellValueMulti.NumberFormat = numberFormat;
            return true;
        }
        catch (Exception ex)
        {
            excelError = new ExcelError(ExcelErrorCode.TypeWrong, ex);
            excelCellValueMulti = null;
            return false;
        }
    }

    public static bool CreateValueDateTime(string value, int numberFormatId, string numberFormat, out ExcelCellValueMulti excelCellValueMulti, out ExcelError excelError)
    {
        excelError = null;

        try
        {
            value = value.Replace('.', ',');
            double valDouble = double.Parse(value);

            // convert the value to date
            DateTime dateTime = DateTime.FromOADate(valDouble);
            excelCellValueMulti = new ExcelCellValueMulti(dateTime);
            excelCellValueMulti.NumberFormatId = numberFormatId;
            excelCellValueMulti.NumberFormat = numberFormat;
            return true;
        }
        catch (Exception ex)
        {
            excelError = new ExcelError(ExcelErrorCode.TypeWrong, ex);
            excelCellValueMulti = null;
            return false;
        }
    }

    public static bool CreateValueTimeOnly(string value, int numberFormatId,  string numberFormat, out ExcelCellValueMulti excelCellValueMulti, out ExcelError excelError)
    {
        excelError = null;

        try
        {
            value = value.Replace('.', ',');
            double valDouble = double.Parse(value);

            // convert the value to date
            DateTime dateTime = DateTime.FromOADate(valDouble);
            TimeOnly timeOnly = TimeOnly.FromDateTime(dateTime);
            excelCellValueMulti = new ExcelCellValueMulti(timeOnly);
            excelCellValueMulti.NumberFormatId = numberFormatId;
            excelCellValueMulti.NumberFormat = numberFormat;
            return true;
        }
        catch (Exception ex)
        {
            excelError = new ExcelError(ExcelErrorCode.TypeWrong, ex);
            excelCellValueMulti = null;
            return false;
        }
    }
}