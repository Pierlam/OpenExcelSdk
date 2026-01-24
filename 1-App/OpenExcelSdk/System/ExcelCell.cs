using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenExcelSdk;


/// <summary>
/// Represents a cell within an Excel worksheet, providing access to its underlying Open XML representation and
/// formatting information.
/// </summary>
/// <remarks>An ExcelCell encapsulates both the logical context (the worksheet it belongs to) and the Open XML
/// elements that define the cell's content and style. Use this type to interact with cell data, retrieve or modify
/// formatting, or access the cell's position within the worksheet. The Cell and CellFormat properties expose the raw
/// Open XML objects for advanced scenarios, such as direct manipulation of cell styles or values.</remarks>
public class ExcelCell
{
    public ExcelCell(ExcelSheet excelSheet, Cell cell)
    {
        ExcelSheet = excelSheet;
        Cell = cell;
    }

    /// <summary>
    /// The sheet where is placed the cell.
    /// </summary>
    public ExcelSheet ExcelSheet { get; set; }

    /// <summary>
    /// Gets the formula expression assigned to the cell, if any.
    /// </summary>
    /// <remarks>Returns an empty string if the cell does not contain a formula. The returned value represents
    /// the formula as a string, without any leading equals sign or formatting.</remarks>
    public string Formula
    {
        get { return Cell.CellFormula?.Text ?? string.Empty; }
    }

    /// <summary>
    /// Open Xml cell object.
    /// </summary>
    public Cell Cell { get; set; }

    /// <summary>
    /// Format of the cell.
    /// OpenXml CellFormat object.
    /// Not null only if the cell has a style. (Cell.StyleIndex).
    /// The data is in the Styles part of the Excel file: excelSheet.ExcelFile.WorkbookPart.WorkbookStylesPart
    /// </summary>
    public CellFormat? CellFormat { get; set; } = null;


    
}