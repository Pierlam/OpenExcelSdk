namespace OpenExcelSdk;

public class ValueBuilder
{
    public static bool CreateValueInteger(string value, string dataFormat, out ExcelCellValueMulti excelCellValueMulti, out ExcelError excelError)
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
        return true;
    }

    public static bool CreateValueDouble(string value, string dataFormat, out ExcelCellValueMulti excelCellValueMulti, out ExcelError excelError)
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
        return true;
    }

    public static bool CreateValueDateOnly(string value, string dataFormat, out ExcelCellValueMulti excelCellValueMulti, out ExcelError excelError)
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
            excelCellValueMulti.DataFormat = dataFormat;
            return true;
        }
        catch (Exception ex)
        {
            excelError = new ExcelError(ExcelErrorCode.TypeWrong, ex);
            excelCellValueMulti = null;
            return false;
        }
    }

    public static bool CreateValueDateTime(string value, string dataFormat, out ExcelCellValueMulti excelCellValueMulti, out ExcelError excelError)
    {
        excelError = null;

        try
        {
            value = value.Replace('.', ',');
            double valDouble = double.Parse(value);

            // convert the value to date
            DateTime dateTime = DateTime.FromOADate(valDouble);
            excelCellValueMulti = new ExcelCellValueMulti(dateTime);
            excelCellValueMulti.DataFormat = dataFormat;
            return true;
        }
        catch (Exception ex)
        {
            excelError = new ExcelError(ExcelErrorCode.TypeWrong, ex);
            excelCellValueMulti = null;
            return false;
        }
    }

    public static bool CreateValueTimeOnly(string value, string dataFormat, out ExcelCellValueMulti excelCellValueMulti, out ExcelError excelError)
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
            excelCellValueMulti.DataFormat = dataFormat;
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