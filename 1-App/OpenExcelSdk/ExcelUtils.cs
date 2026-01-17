using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenExcelSdk;

public class ExcelUtils
{
    /// <summary>
    /// Get the style/CellFormat of the cell, if it has one.
    /// It's an OpenXml object.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="excelCell"></param>
    /// <returns></returns>
    public static CellFormat GetCellFormat(ExcelSheet excelSheet, ExcelCell excelCell)
    {
        if (excelCell.Cell.StyleIndex == null)
            // no style, no cell format
            return null;

        var stylesPart = excelSheet.ExcelFile.WorkbookPart.WorkbookStylesPart;
        return (CellFormat)stylesPart.Stylesheet.CellFormats.ElementAt((int)excelCell.Cell.StyleIndex.Value);
    }

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
    /// Return column and row index from a cell address.
    /// e.g. B2 -> colIdx=2, rowIdx=2
    /// </summary>
    /// <param name="cellReference"></param>
    /// <param name="colIdx"></param>
    /// <param name="rowIdx"></param>
    /// <returns></returns>
    public static bool GetColumnAndRowIndex(string cellReference, out int colIdx, out int rowIdx)
    {
        colIdx = 0;
        rowIdx = 0;

        if (string.IsNullOrWhiteSpace(cellReference))
        {
            return false;
        }

        cellReference= cellReference.Trim();  
        
        colIdx = GetColumnIndex(cellReference);
        if(colIdx == 0) return false;

        rowIdx = GetRowIndex(cellReference);
        if (rowIdx == 0) return false;

        // max col: XFD, max row: 1048576
        return CheckMaxColAndRowValue(colIdx, rowIdx);
    }


    /// <summary>
    /// Get the column index.
    /// exp: B2 -> return 2.
    /// </summary>
    /// <param name="cellReference"></param>
    /// <returns></returns>
    public static int GetColumnIndex(string cellReference)
    {
        if (string.IsNullOrWhiteSpace(cellReference)) return 0;

        string columnAddress = string.Empty;
        int i = 0;
        while (true)
        {
            if (i >= cellReference.Length) break;

            if (char.IsLetter(cellReference[i]))
            {
                columnAddress += cellReference[i];
                i++;
                continue;
            }
            break;
        }

        if (columnAddress == string.Empty) return 0;

        columnAddress= columnAddress.ToUpper();

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
    /// <param name="cellReference"></param>
    /// <returns></returns>
    public static int GetRowIndex(string cellReference)
    {
        if (string.IsNullOrWhiteSpace(cellReference)) return 0;

        int i = 0;

        // scan letters part, e.g. A
        while (true)
        {
            if (i >= cellReference.Length) break;
            if (char.IsLetter(cellReference[i]))
            {
                i++;
                continue;
            }
            break;
        }

        // scan digits part, e.g. 12
        string rowStr = string.Empty;
        while (true)
        {
            if (i >= cellReference.Length) break;
            if (char.IsDigit(cellReference[i]))
            {
                rowStr += cellReference[i];
                i++;
                continue;
            }
            break;
        }

        // next char are not digit
        if (i < cellReference.Length) return 0;

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