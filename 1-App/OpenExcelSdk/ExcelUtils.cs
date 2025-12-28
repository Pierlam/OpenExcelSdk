namespace OpenExcelSdk;

public class ExcelUtils
{
    /// <summary>
    /// Convert to a standard excel address.
    /// exp: 1,1 -> A1
    /// </summary>
    /// <param name="col"></param>
    /// <param name="rowIndex"></param>
    /// <returns></returns>
    public static string ConvertAddress(int colIndex, int rowIndex)
    {
        if (colIndex < 1) return string.Empty;
        if (rowIndex < 1) return string.Empty;

        return GetColumnName(colIndex) + rowIndex.ToString();
    }

    /// <summary>
    /// Get the column index.
    /// exp: B2 -> return 2.
    /// </summary>
    /// <param name="addressName"></param>
    /// <returns></returns>
    public static int GetColIndex(string addressName)
    {
        if (string.IsNullOrWhiteSpace(addressName)) return 0;
        string columnAddress = string.Empty;
        int i = 0;
        while (true)
        {
            if (i >= addressName.Length) break;

            if (char.IsLetter(addressName[i]))
            {
                columnAddress += addressName[i];
                i++;
                continue;
            }
            break;
        }

        if (columnAddress == string.Empty) return 0;

        // convert the col to an int
        int columnNumber = 0;
        foreach (char c in columnAddress)
        {
            columnNumber = columnNumber * 26 + (c - 'A' + 1);
        }

        return columnNumber;
    }

    /// <summary>
    /// Get the row value from a cell address.
    /// </summary>
    /// <param name="addressName"></param>
    /// <returns></returns>
    public static int GetRowIndex(string addressName)
    {
        if (string.IsNullOrWhiteSpace(addressName)) return 0;
        int i = 0;
        while (true)
        {
            if (i >= addressName.Length) break;
            if (char.IsLetter(addressName[i]))
            {
                i++;
                continue;
            }
            break;
        }
        string rowStr = string.Empty;
        while (true)
        {
            if (i >= addressName.Length) break;
            if (char.IsDigit(addressName[i]))
            {
                rowStr += addressName[i];
                i++;
                continue;
            }
            break;
        }

        int row = 0;
        if (!int.TryParse(rowStr, out row)) return 0;
        return row;
    }

    /// <summary>
    /// return the column name of the col index.
    /// exp: 1 -> A
    /// </summary>
    /// <param name="index"></param>
    /// <returns></returns>
    public static string GetColumnName(int index)
    {
        if (index < 1) return String.Empty;

        index--;
        const string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

        var value = "";

        if (index >= letters.Length)
            value += letters[index / letters.Length - 1];

        value += letters[index % letters.Length];

        return value;
    }

    public static bool CheckMaxColAndRowValue(int colIndex, int rowIndex)
    {
        // XFD
        if (colIndex > 16384) return false;
        // 1,048,576
        if (rowIndex > 1048576) return false;

        return true;
    }
}