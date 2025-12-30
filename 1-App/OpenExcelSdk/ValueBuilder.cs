namespace OpenExcelSdk;

/// <summary>
/// To build excelCellValue object.
/// type and value of a cell.
/// </summary>
public class ValueBuilder
{
    public static ExcelCellValue CreateValue(ExcelCell excelCell, ExcelCellType cellType, string value, int numberFormatId, string numberFormat)
    {
        ExcelCellValue excelCellValue;

        if (string.IsNullOrEmpty(value))
        {
            excelCellValue = new ExcelCellValue();
            excelCellValue.CellType = cellType;
            excelCellValue.IsEmpty = true;
            return excelCellValue;
        }

        if (cellType == ExcelCellType.Integer)
        {
            excelCellValue = ValueBuilder.CreateValueInteger(value, (int)numberFormatId, numberFormat);
            if (excelCellValue == null) return null;
            excelCellValue.Formula = excelCell.Cell.CellFormula?.Text;
            return excelCellValue;
        }

        if (cellType == ExcelCellType.Double)
        {
            excelCellValue = ValueBuilder.CreateValueDouble(value, (int)numberFormatId, numberFormat);
            if (excelCellValue == null) return null;
            excelCellValue.Formula = excelCell.Cell.CellFormula?.Text;
            return excelCellValue;
        }

        if (cellType == ExcelCellType.DateOnly)
        {
            excelCellValue = ValueBuilder.CreateValueDateOnly(value, (int)numberFormatId, numberFormat);
            if (excelCellValue == null) return null;
            excelCellValue.Formula = excelCell.Cell.CellFormula?.Text;
            return excelCellValue;
        }

        if (cellType == ExcelCellType.DateTime)
        {
            excelCellValue = ValueBuilder.CreateValueDateTime(value, (int)numberFormatId, numberFormat);
            if (excelCellValue == null) return null;
            excelCellValue.Formula = excelCell.Cell.CellFormula?.Text;
            return excelCellValue;
        }

        if (cellType == ExcelCellType.TimeOnly)
        {
            excelCellValue = ValueBuilder.CreateValueTimeOnly(value, (int)numberFormatId, numberFormat);
            if (excelCellValue == null) return null;
            excelCellValue.Formula = excelCell.Cell.CellFormula?.Text;
            return excelCellValue;
        }

        excelCellValue = new ExcelCellValue();
        excelCellValue.CellType = cellType;
        return excelCellValue;
    }

    public static ExcelCellValue CreateValueInteger(string value, int numberFormatId,  string numberFormat)
    {
        bool res = int.TryParse(value, out int valInt);
        if (!res)
            throw ExcelException.Create("CreateValueInteger", ExcelErrorCode.TypeWrong, value);

        var excelCellValue = new ExcelCellValue(valInt);
        excelCellValue.NumberFormatId = numberFormatId;
        excelCellValue.NumberFormat = numberFormat;
        return excelCellValue;
    }

    public static ExcelCellValue CreateValueDouble(string value, int numberFormatId, string numberFormat)
    {
        // cultureInfo prb: replace . by ,
        value = value.Replace('.', ',');
        bool res = double.TryParse(value, out double valDouble);
        if (!res)
            throw ExcelException.Create("CreateValueDouble", ExcelErrorCode.TypeWrong, value);

        var excelCellValue = new ExcelCellValue(valDouble);
        excelCellValue.NumberFormatId = numberFormatId;
        excelCellValue.NumberFormat = numberFormat;
        return excelCellValue;
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
            var excelCellValue = new ExcelCellValue(dateOnly);
            excelCellValue.NumberFormatId = numberFormatId;
            excelCellValue.NumberFormat = numberFormat;
            return excelCellValue;
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
            var excelCellValue = new ExcelCellValue(dateTime);
            excelCellValue.NumberFormatId = numberFormatId;
            excelCellValue.NumberFormat = numberFormat;
            return excelCellValue;
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
            var excelCellValue = new ExcelCellValue(timeOnly);
            excelCellValue.NumberFormatId = numberFormatId;
            excelCellValue.NumberFormat = numberFormat;
            return excelCellValue;
        }
        catch (Exception ex)
        {
            throw ExcelException.Create("CreateValueTimeOnly", ExcelErrorCode.TypeWrong, value, ex);
        }
    }
}