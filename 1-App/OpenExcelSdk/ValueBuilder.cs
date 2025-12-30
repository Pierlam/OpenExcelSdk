namespace OpenExcelSdk;

/// <summary>
/// To build ExcelCellValueMulti object.
/// type and value of a cell.
/// </summary>
public class ValueBuilder
{
    public static ExcelCellValue CreateValue(ExcelCell excelCell, ExcelCellType cellType, string value, int numberFormatId, string numberFormat)
    {
        ExcelCellValue excelCellValueMulti;

        if (string.IsNullOrEmpty(value))
        {
            excelCellValueMulti = new ExcelCellValue();
            excelCellValueMulti.CellType = cellType;
            excelCellValueMulti.IsEmpty = true;
            return excelCellValueMulti;
        }

        if (cellType == ExcelCellType.Integer)
        {
            excelCellValueMulti = ValueBuilder.CreateValueInteger(value, (int)numberFormatId, numberFormat);
            if (excelCellValueMulti == null) return null;
            excelCellValueMulti.Formula = excelCell.Cell.CellFormula?.Text;
            return excelCellValueMulti;
        }

        if (cellType == ExcelCellType.Double)
        {
            excelCellValueMulti = ValueBuilder.CreateValueDouble(value, (int)numberFormatId, numberFormat);
            if (excelCellValueMulti == null) return null;
            excelCellValueMulti.Formula = excelCell.Cell.CellFormula?.Text;
            return excelCellValueMulti;
        }

        if (cellType == ExcelCellType.DateOnly)
        {
            excelCellValueMulti = ValueBuilder.CreateValueDateOnly(value, (int)numberFormatId, numberFormat);
            if (excelCellValueMulti == null) return null;
            excelCellValueMulti.Formula = excelCell.Cell.CellFormula?.Text;
            return excelCellValueMulti;
        }

        if (cellType == ExcelCellType.DateTime)
        {
            excelCellValueMulti = ValueBuilder.CreateValueDateTime(value, (int)numberFormatId, numberFormat);
            if (excelCellValueMulti == null) return null;
            excelCellValueMulti.Formula = excelCell.Cell.CellFormula?.Text;
            return excelCellValueMulti;
        }

        if (cellType == ExcelCellType.TimeOnly)
        {
            excelCellValueMulti = ValueBuilder.CreateValueTimeOnly(value, (int)numberFormatId, numberFormat);
            if (excelCellValueMulti == null) return null;
            excelCellValueMulti.Formula = excelCell.Cell.CellFormula?.Text;
            return excelCellValueMulti;
        }

        excelCellValueMulti = new ExcelCellValue();
        excelCellValueMulti.CellType = cellType;
        return excelCellValueMulti;
    }

    public static ExcelCellValue CreateValueInteger(string value, int numberFormatId,  string numberFormat)
    {
        bool res = int.TryParse(value, out int valInt);
        if (!res)
            throw ExcelException.Create("CreateValueInteger", ExcelErrorCode.TypeWrong, value);

        var excelCellValueMulti = new ExcelCellValue(valInt);
        excelCellValueMulti.NumberFormatId = numberFormatId;
        excelCellValueMulti.NumberFormat = numberFormat;
        return excelCellValueMulti;
    }

    public static ExcelCellValue CreateValueDouble(string value, int numberFormatId, string numberFormat)
    {
        // cultureInfo prb: replace . by ,
        value = value.Replace('.', ',');
        bool res = double.TryParse(value, out double valDouble);
        if (!res)
            throw ExcelException.Create("CreateValueDouble", ExcelErrorCode.TypeWrong, value);

        var excelCellValueMulti = new ExcelCellValue(valDouble);
        excelCellValueMulti.NumberFormatId = numberFormatId;
        excelCellValueMulti.NumberFormat = numberFormat;
        return excelCellValueMulti;
    }

    public static ExcelCellValue CreateValueDateOnly(string value, int numberFormatId, string numberFormat)
    {
        try
        {
            value = value.Replace('.', ',');
            double valDouble = double.Parse(value);

            // convert the value to date
            DateTime dateTime = DateTime.FromOADate(valDouble);
            DateOnly dateOnly = DateOnly.FromDateTime(dateTime);
            var excelCellValueMulti = new ExcelCellValue(dateOnly);
            excelCellValueMulti.NumberFormatId = numberFormatId;
            excelCellValueMulti.NumberFormat = numberFormat;
            return excelCellValueMulti;
        }
        catch (Exception ex)
        {
            throw ExcelException.Create("CreateValueDateOnly", ExcelErrorCode.TypeWrong, value, ex);
        }
    }

    public static ExcelCellValue CreateValueDateTime(string value, int numberFormatId, string numberFormat)
    {
        try
        {
            value = value.Replace('.', ',');
            double valDouble = double.Parse(value);

            // convert the value to date
            DateTime dateTime = DateTime.FromOADate(valDouble);
            var excelCellValueMulti = new ExcelCellValue(dateTime);
            excelCellValueMulti.NumberFormatId = numberFormatId;
            excelCellValueMulti.NumberFormat = numberFormat;
            return excelCellValueMulti;
        }
        catch (Exception ex)
        {
            throw ExcelException.Create("CreateValueDateTime", ExcelErrorCode.TypeWrong, value, ex);
        }
    }

    public static ExcelCellValue CreateValueTimeOnly(string value, int numberFormatId,  string numberFormat)
    {
        try
        {
            value = value.Replace('.', ',');
            double valDouble = double.Parse(value);

            // convert the value to date
            DateTime dateTime = DateTime.FromOADate(valDouble);
            TimeOnly timeOnly = TimeOnly.FromDateTime(dateTime);
            var excelCellValueMulti = new ExcelCellValue(timeOnly);
            excelCellValueMulti.NumberFormatId = numberFormatId;
            excelCellValueMulti.NumberFormat = numberFormat;
            return excelCellValueMulti;
        }
        catch (Exception ex)
        {
            throw ExcelException.Create("CreateValueTimeOnly", ExcelErrorCode.TypeWrong, value, ex);
        }
    }
}