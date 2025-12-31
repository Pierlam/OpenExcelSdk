using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenExcelSdk;

/// <summary>
/// Main class to process Excel files.
/// Open/close, create excel file.
/// Get/Create sheet. Get/Create cell.
/// Get/set cell value.
/// and more...
/// </summary>
public class ExcelProcessor : ExcelProcessorBase
{
    #region Open/Close Create Excel file

    /// <summary>
    /// Open an existing Excel file.
    /// </summary>
    /// <param name="fileName"></param>
    /// <returns></returns>
    public ExcelFile OpenExcelFile(string fileName)
    {
        if (string.IsNullOrWhiteSpace(fileName))
            throw ExcelException.Create("OpenExcel", ExcelErrorCode.FilenameNull);

        if (!File.Exists(fileName))
            throw ExcelException.Create("OpenExcel", ExcelErrorCode.FileNotFound, fileName);

        try
        {
            // Open the document for editing.
            SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, true);
            return new ExcelFile(fileName, document);
        }
        catch (Exception ex)
        {
            throw ExcelException.Create("OpenExcel", ExcelErrorCode.UnableOpenFile, fileName, ex);
        }
    }

    /// <summary>
    /// Close an open excel file.
    /// </summary>
    /// <param name="excelFile"></param>
    /// <returns></returns>
    public void CloseExcelFile(ExcelFile excelFile)
    {
        try
        {
            excelFile.SpreadsheetDocument.Dispose();
            excelFile.SpreadsheetDocument = null;
        }
        catch (Exception ex)
        {
            throw ExcelException.Create("CloseExcel", ExcelErrorCode.UnableOpenFile, excelFile.Filename, ex);
        }
    }

    /// <summary>
    /// Create a new excel file with one sheet.
    /// The filename should not exists.
    /// exp: "C:\Files\MyExcel.xlsx"
    /// </summary>
    /// <param name="fileName"></param>
    /// <returns></returns>
    public ExcelFile CreateExcelFile(string fileName)
    {
        return CreateExcelFile(fileName, Definitions.DefaultFirstSheetName);
    }

    /// <summary>
    /// Create a new excel file with one sheet. Provide the sheet name.
    /// The filename should not exists.
    /// exp: "C:\Files\MyExcel.xlsx"
    ///
    /// https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/structure-of-a-spreadsheetml-document?tabs=cs
    /// </summary>
    /// <param name="fileName"></param>
    public ExcelFile CreateExcelFile(string fileName, string sheetName)
    {
        if (string.IsNullOrWhiteSpace(fileName))
            throw ExcelException.Create("CreateExcelFile", ExcelErrorCode.FilenameNull);

        if (File.Exists(fileName))
            throw ExcelException.Create("CreateExcelFile", ExcelErrorCode.FileAlreadyExists, fileName);

        if (string.IsNullOrWhiteSpace(sheetName))
            throw ExcelException.Create("CreateExcelFile", ExcelErrorCode.SheetnameNull);

        try
        {
            // Create a spreadsheet document by supplying the filepath.
            // By default, AutoSave = true, Editable = true, and Type = xlsx.
            SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook);

            // Add a WorkbookPart to the document.
            WorkbookPart workbookPart = spreadsheetDocument.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            // Add a WorksheetPart to the WorkbookPart.
            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            // Add Sheets to the Workbook.
            Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());

            // Append a new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = sheetName };
            sheets.Append(sheet);

            return new ExcelFile(fileName, spreadsheetDocument);
        }
        catch (Exception ex)
        {
            throw ExcelException.Create("CreateExcelFile", ExcelErrorCode.UnableCreateFile, fileName, ex);
        }
    }

    #endregion Open/Close Create Excel file

    #region Get/Create sheet

    /// <summary>
    /// Create a new sheet in the Excel file.
    /// Provide the sheet name, should be unique.
    /// The sheet will be added after the last existing one.
    /// </summary>
    /// <param name="excelFile"></param>
    /// <param name="sheetName"></param>
    /// <returns></returns>
    public ExcelSheet CreateSheet(ExcelFile excelFile, string sheetName)
    {
        if (string.IsNullOrWhiteSpace(sheetName))
            throw ExcelException.Create("CreateSheet", ExcelErrorCode.SheetnameNull);

        if (excelFile.WorkbookPart.Workbook == null)
            // Add Sheets to the Workbook.
            excelFile.WorkbookPart.Workbook.AppendChild(new Sheets());

        var listSheets = excelFile.WorkbookPart.Workbook.Sheets.Elements<Sheet>();

        // Find the sheet with the matching name
        var sheet = listSheets.FirstOrDefault(s => s.Name != null && s.Name.Value.Equals(sheetName, StringComparison.InvariantCultureIgnoreCase));
        if (sheet != null)
            throw ExcelException.Create("CreateSheet", ExcelErrorCode.UnableCreateSheet);

        // get the last id
        uint sheetId = 1;
        Sheets sheets = excelFile.WorkbookPart.Workbook.Sheets;
        if (sheets.Elements<Sheet>().Any())
        {
            sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
        }

        // Add a WorksheetPart to the WorkbookPart
        WorksheetPart worksheetPart = excelFile.WorkbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(new SheetData());

        // Get the relationship ID of the worksheetPart from the workbookPart
        string relId = excelFile.WorkbookPart.GetIdOfPart(worksheetPart);

        // Append a new worksheet and associate it with the workbook.
        sheet = new Sheet() { Id = relId, SheetId = sheetId, Name = sheetName };
        sheets.Append(sheet);

        var excelSheet = new ExcelSheet(excelFile, sheet);
        excelSheet.Index = (int)sheetId;
        excelSheet.Name = sheet.Name;
        return excelSheet;
    }

    /// <summary>
    /// Get the first sheet of the excel file.
    /// </summary>
    /// <param name="excelFile"></param>
    /// <param name="index"></param>
    /// <returns></returns>
    public ExcelSheet GetFirstSheet(ExcelFile excelFile, int index)
    {
        return GetSheetAt(excelFile, 0);
    }

    /// <summary>
    /// Get a sheet of the excel file by index base0.
    /// </summary>
    /// <param name="excelFile"></param>
    /// <param name="index"></param>
    /// <returns></returns>
    public ExcelSheet GetSheetAt(ExcelFile excelFile, int index)
    {
        if (index < 0)
            throw ExcelException.Create("GetSheetAt", ExcelErrorCode.IndexMustBePositive, index.ToString());

        if (excelFile == null)
            throw ExcelException.Create("GetSheetAt", ExcelErrorCode.FilenameNull);

        try
        {
            Sheet? sheet = excelFile.WorkbookPart?.Workbook?.GetFirstChild<Sheets>()?.Elements<Sheet>()?.ElementAt<Sheet>(index);

            if (sheet == null)
                throw ExcelException.Create("GetSheetAt", ExcelErrorCode.IndexWrong, index.ToString());

            var excelSheet = new ExcelSheet(excelFile, sheet);
            excelSheet.Index = index;
            excelSheet.Name = sheet.Name;
            return excelSheet;
        }
        catch (Exception ex)
        {
            throw ExcelException.Create("GetSheetAt", ExcelErrorCode.UnableGetSheet, index.ToString(), ex);
        }
    }

    /// <summary>
    /// Get a sheet of the excel file by the name.
    /// </summary>
    /// <param name="excelFile"></param>
    /// <param name="sheetName"></param>
    /// <returns></returns>
    public ExcelSheet GetSheetByName(ExcelFile excelFile, string sheetName)
    {
        if (string.IsNullOrWhiteSpace(sheetName))
            throw ExcelException.Create("GetSheetByName", ExcelErrorCode.SheetnameNull);

        try
        {
            // Get the Sheets collection
            var sheets = excelFile.WorkbookPart.Workbook.Sheets.Elements<Sheet>();

            // Find the sheet with the matching name
            var sheet = sheets.FirstOrDefault(s => s.Name != null && s.Name.Value.Equals(sheetName, StringComparison.OrdinalIgnoreCase));
            if (sheet == null)
                // no sheet with this name, not an error
                return null;

            var excelSheet = new ExcelSheet(excelFile, sheet);
            // excelSheet.Index = index;  TO_IMPLEMENT!
            excelSheet.Name = sheet.Name;

            return excelSheet;
        }
        catch (Exception ex)
        {
            throw ExcelException.Create("GetSheetByName", ExcelErrorCode.UnableGetSheet, sheetName, ex);
        }
    }

    #endregion Get/Create sheet

    #region Get row

    /// <summary>
    /// Get a row from the sheet  by index base0.
    /// If the row doest not exists, it's not an error.
    /// Error occurs only if the access to the row fails.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="rowIndex"></param>
    /// <returns></returns>
    public ExcelRow GetRowAt(ExcelSheet excelSheet, int rowIndex)
    {
        try
        {
            var rows = excelSheet.Worksheet.Descendants<Row>();
            if (!rows.Any())
                // row doest not exists, it's not an error
                return null;

            if (rowIndex < 0 || rowIndex > rows.Count())
                throw ExcelException.Create("GetRowAt", ExcelErrorCode.IndexWrong, rowIndex.ToString());

            Row row = rows.ElementAt(rowIndex);
            return new ExcelRow(row);
        }
        catch (Exception ex)
        {
            throw ExcelException.Create("GetRowAt", ExcelErrorCode.UnableGetRow, rowIndex.ToString(), ex);
        }
    }

    /// <summary>
    /// Return the last row index, base1.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <returns></returns>
    public int GetLastRowIndex(ExcelSheet excelSheet)
    {
        if (!excelSheet.Rows.Any()) return 0;

        // Get the last row index
        //return rows.Max(r => r.RowIndex.Value);
        return excelSheet.Rows.Count();
    }

    #endregion Get row



    #region Get Cell at

    /// <summary>
    /// Get a cell in the sheet by col and row index, base1.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="colIdx"></param>
    /// <param name="rowIdx"></param>
    /// <returns></returns>
    public ExcelCell GetCellAt(ExcelSheet excelSheet, int colIdx, int rowIdx)
    {
        // convert the col and the rox to an excel address
        return GetCellAt(excelSheet, ExcelUtils.ConvertAddress(colIdx, rowIdx));
    }

    /// <summary>
    /// Get a cell in the sheet by the address name. exp: A1
    ///
    /// If the cell does not exists, return a  null ExcelCell without error.
    /// If the access to the cell fails, then return an error.
    /// https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/how-to-retrieve-the-values-of-cells-in-a-spreadsheet?tabs=cs-0%2Ccs-2%2Ccs-3%2Ccs-4%2Ccs-5%2Ccs-6%2Ccs-7%2Ccs-8%2Ccs-9%2Ccs-10%2Ccs-11%2Ccs
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="cellReference"></param>
    /// <returns></returns>
    public ExcelCell GetCellAt(ExcelSheet excelSheet, string cellReference)
    {
        if (excelSheet == null)
            throw ExcelException.Create("GetCellAt", ExcelErrorCode.ObjectNull);

        try
        {
            Cell? cell = excelSheet.Worksheet?.Descendants<Cell>()?.Where(c => c.CellReference == cellReference).FirstOrDefault();
            if (cell == null)
                // no cell found, not an error
                return null;

            var excelCell = new ExcelCell(excelSheet, cell);

            // get the style of the cell
            excelCell.CellFormat = GetCellFormat(excelSheet, excelCell);

            return excelCell;
        }
        catch (Exception ex)
        {
            throw ExcelException.Create("GetCellAt", ExcelErrorCode.UnableGetCell, cellReference, ex);
        }
    }

    #endregion Get Cell at

    #region Get CellType

    /// <summary>
    /// Get the type of the cell value.
    /// If the cell is empty/blank, in some cases the type will be Undefined.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="cellReference"></param>
    /// <returns></returns>
    public ExcelCellType GetCellType(ExcelSheet excelSheet, string cellReference)
    {
        var excelCellValue = GetCellValue(excelSheet, cellReference);
        if (excelCellValue == null) return ExcelCellType.Undefined;
        return excelCellValue.CellType;
    }

    /// <summary>
    /// Get the type of the cell value.
    /// If the cell is empty/blank, in some cases the type will be Undefined.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="colIdx"></param>
    /// <param name="rowIdx"></param>
    /// <returns></returns>
    public ExcelCellType GetCellType(ExcelSheet excelSheet, int colIdx, int rowIdx)
    {
        var excelCellValue = GetCellValue(excelSheet, colIdx, rowIdx);
        if (excelCellValue == null) return ExcelCellType.Undefined;
        return excelCellValue.CellType;
    }

    #endregion Get CellType

    #region Get CellValue as

    //public string GetCellValueAsString(ExcelSheet excelSheet, ExcelCell excelCell)
    //public string GetCellValueAsDouble(ExcelSheet excelSheet, ExcelCell excelCell)

    #endregion Get CellValue as

    #region Get CellValue

    /// <summary>
    /// Get the type, the value and the data format of the cell.
    /// If the cell is empty/blank, in some cases the type will be Undefined.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="cellReference"></param>
    /// <returns></returns>
    public ExcelCellValue GetCellValue(ExcelSheet excelSheet, string cellReference)
    {
        ExcelCell excelCell = GetCellAt(excelSheet, cellReference);
        if (excelCell == null) return null;

        return GetCellValue(excelSheet, excelCell);
    }

    /// <summary>
    /// Geth the type of the cell. If the cell is empty/blank, in some cases the type will be Undefined.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="colIdx"></param>
    /// <param name="rowIdx"></param>
    /// <returns></returns>
    public ExcelCellValue GetCellValue(ExcelSheet excelSheet, int colIdx, int rowIdx)
    {
        ExcelCell excelCell = GetCellAt(excelSheet, colIdx, rowIdx);
        if (excelCell == null) return null;

        return GetCellValue(excelSheet, excelCell);
    }

    #endregion Get CellValue

    #region Create cell

    /// <summary>
    /// Given an address name e.g. A3, and a WorksheetPart, inserts a cell into the worksheet.
    /// If the cell already exists, returns it.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="cellReference"></param>
    /// <returns></returns>
    public ExcelCell CreateCell(ExcelSheet excelSheet, string cellReference)
    {
        string colName = ExcelUtils.GetColumnName(ExcelUtils.GetColumnIndex(cellReference));

        return CreateCell(excelSheet, colName, (uint)ExcelUtils.GetRowIndex(cellReference));
    }

    /// <summary>
    /// Given a column name, a row index, and a WorksheetPart, inserts a cell into the worksheet.
    /// If the cell already exists, returns it.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="colIdx"></param>
    /// <param name="rowIdx"></param>
    /// <returns></returns>
    public ExcelCell CreateCell(ExcelSheet excelSheet, int colIdx, int rowIdx)
    {
        string colName = ExcelUtils.GetColumnName(colIdx);
        return CreateCell(excelSheet, colName, (uint)rowIdx);
    }

    #endregion Create cell

    #region Remove cell

    /// <summary>
    /// Remove a cell in the sheet by col and row index, start at index 1.
    /// If there is not cell at the address, no eror will occur.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="cellReference"></param>
    /// <returns></returns>
    public bool RemoveCell(ExcelSheet excelSheet, string cellReference)
    {
        return RemoveCell(excelSheet, ExcelUtils.GetColumnIndex(cellReference), ExcelUtils.GetRowIndex(cellReference));
    }

    /// <summary>
    /// Remove a cell in the sheet by the address name. exp: A1
    /// If there is not cell at the address, no eror will occur.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="colIdx"></param>
    /// <param name="rowIdx"></param>
    /// <returns></returns>
    public bool RemoveCell(ExcelSheet excelSheet, int colIdx, int rowIdx)
    {
        ExcelCell excelCell = GetCellAt(excelSheet, colIdx, rowIdx);
        if (excelCell == null || excelCell.Cell == null)
        {
            // no cell at this address, not an error
            return true;
        }
        try
        {
            excelCell.Cell.Remove();
            return true;
        }
        catch (Exception ex)
        {
            throw ExcelException.Create("RemoveCell", ExcelErrorCode.UnableRemoveCell, ExcelUtils.ConvertAddress(colIdx, rowIdx), ex);
        }
    }

    #endregion Remove cell

    #region Set cell value

    /// <summary>
    /// Empty/Clear a cell value.
    /// Keep the format: Alignement colors, border, ...
    /// If the cell contains a formula, remove it.
    /// It the cell is null, do nothing.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="cellReference"></param>
    /// <returns></returns>
    public bool SetCellValueEmpty(ExcelSheet excelSheet, string cellReference)
    {
        ExcelCell excelCell = GetCellAt(excelSheet, cellReference);
        if (excelCell == null || excelCell.Cell == null)
        {
            // no cell at this address
            return false;
        }

        excelCell.Cell.CellValue = new CellValue(string.Empty);

        // remove formula if it's there
        _styleMgr.RemoveFormula(excelSheet, excelCell);
        return true;
    }

    /// <summary>
    /// Set a string value in the cell.
    /// If the cell does not exist, it will be created.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="cellReference"></param>
    /// <param name="value"></param>
    /// <returns></returns>
    public bool SetCellValue(ExcelSheet excelSheet, string cellReference, string value)
    {
        return SetCellValue(excelSheet, ExcelUtils.GetColumnIndex(cellReference), ExcelUtils.GetRowIndex(cellReference), value);
    }

    /// <summary>
    /// Set an int value in the cell.
    /// If the cell does not exist, it will be created.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="cellReference"></param>
    /// <param name="value"></param>
    /// <returns></returns>
    public bool SetCellValue(ExcelSheet excelSheet, string cellReference, int value)
    {
        return SetCellValue(excelSheet, ExcelUtils.GetColumnIndex(cellReference), ExcelUtils.GetRowIndex(cellReference), value);
    }

    /// <summary>
    /// Set a double value in the cell.
    /// If the cell does not exist, it will be created.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="cellReference"></param>
    /// <param name="value"></param>
    /// <returns></returns>
    public bool SetCellValue(ExcelSheet excelSheet, string cellReference, double value)
    {
        return SetCellValue(excelSheet, ExcelUtils.GetColumnIndex(cellReference), ExcelUtils.GetRowIndex(cellReference), value);
    }

    /// <summary>
    /// Set a DateOnly value in the cell.
    /// If the cell does not exist, it will be created.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="cellReference"></param>
    /// <param name="value"></param>
    /// <param name="numberFormat"></param>
    /// <returns></returns>
    public bool SetCellValue(ExcelSheet excelSheet, string cellReference, DateTime value, string numberFormat)
    {
        return SetCellValue(excelSheet, ExcelUtils.GetColumnIndex(cellReference), ExcelUtils.GetRowIndex(cellReference), value, numberFormat);
    }

    /// <summary>
    /// Set a TimeOnly value in the cell.
    /// If the cell does not exist, it will be created.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="cellReference"></param>
    /// <param name="value"></param>
    /// <param name="numberFormat"></param>
    /// <returns></returns>
    public bool SetCellValue(ExcelSheet excelSheet, string cellReference, TimeOnly value, string numberFormat)
    {
        return SetCellValue(excelSheet, ExcelUtils.GetColumnIndex(cellReference), ExcelUtils.GetRowIndex(cellReference), value, numberFormat);
    }

    /// <summary>
    /// Set a DateOnly value in the cell.
    /// If the cell does not exist, it will be created.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="cellReference"></param>
    /// <param name="value"></param>
    /// <param name="numberFormat"></param>
    /// <returns></returns>
    public bool SetCellValue(ExcelSheet excelSheet, string cellReference, DateOnly value, string numberFormat)
    {
        return SetCellValue(excelSheet, ExcelUtils.GetColumnIndex(cellReference), ExcelUtils.GetRowIndex(cellReference), value, numberFormat);
    }

    /// <summary>
    /// Empty/Clear a cell value.
    /// Keep the formating.
    /// If the cell contains a formula, remove it.
    /// It the cell is null, do nothing.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="colIdx"></param>
    /// <param name="rowIdx"></param>
    /// <returns></returns>
    public bool SetCellValueEmpty(ExcelSheet excelSheet, int colIdx, int rowIdx)
    {
        ExcelCell excelCell = GetCellAt(excelSheet, colIdx, rowIdx);

        if (excelCell == null || excelCell.Cell == null) return false;

        excelCell.Cell.CellValue = new CellValue(string.Empty);

        // remove formula if it's there
        _styleMgr.RemoveFormula(excelSheet, excelCell);
        return true;
    }

    /// <summary>
    /// Set a string value in the cell.
    /// If the cell does not exist, it will be created.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="colIdx"></param>
    /// <param name="rowIdx"></param>
    /// <param name="value"></param>
    /// <returns></returns>
    public bool SetCellValue(ExcelSheet excelSheet, int colIdx, int rowIdx, string value)
    {
        string colName = ExcelUtils.GetColumnName(colIdx);
        ExcelCell excelCell = CreateCell(excelSheet, colName, (uint)rowIdx);
        return SetCellValue(excelSheet, excelCell, value);
    }

    /// <summary>
    /// Set an int value in the cell.
    /// If the cell does not exist, it will be created.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="colIdx"></param>
    /// <param name="rowIdx"></param>
    /// <param name="value"></param>
    /// <returns></returns>
    public bool SetCellValue(ExcelSheet excelSheet, int colIdx, int rowIdx, int value)
    {
        string colName = ExcelUtils.GetColumnName(colIdx);
        ExcelCell excelCell = CreateCell(excelSheet, colName, (uint)rowIdx);
        return SetCellValue(excelSheet, excelCell, value);
    }

    /// <summary>
    /// Set a double value in the cell.
    /// If the cell does not exist, it will be created.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="colIdx"></param>
    /// <param name="rowIdx"></param>
    /// <param name="value"></param>
    /// <returns></returns>
    public bool SetCellValue(ExcelSheet excelSheet, int colIdx, int rowIdx, double value)
    {
        string colName = ExcelUtils.GetColumnName(colIdx);
        ExcelCell excelCell = CreateCell(excelSheet, colName, (uint)rowIdx);
        return SetCellValue(excelSheet, excelCell, value);
    }

    /// <summary>
    /// Set a DateOnly value in the cell.
    /// If the cell does not exist, it will be created.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="colIdx"></param>
    /// <param name="rowIdx"></param>
    /// <param name="value"></param>
    /// <param name="numberFormat"></param>
    /// <returns></returns>
    public bool SetCellValue(ExcelSheet excelSheet, int colIdx, int rowIdx, DateOnly value, string numberFormat)
    {
        string colName = ExcelUtils.GetColumnName(colIdx);
        ExcelCell excelCell = CreateCell(excelSheet, colName, (uint)rowIdx);
        return SetCellValue(excelSheet, excelCell, value, numberFormat);
    }

    /// <summary>
    /// Set a DateTime value in the cell.
    /// If the cell does not exist, it will be created.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="colIdx"></param>
    /// <param name="rowIdx"></param>
    /// <param name="value"></param>
    /// <param name="numberFormat"></param>
    /// <returns></returns>
    public bool SetCellValue(ExcelSheet excelSheet, int colIdx, int rowIdx, DateTime value, string numberFormat)
    {
        string colName = ExcelUtils.GetColumnName(colIdx);
        ExcelCell excelCell = CreateCell(excelSheet, colName, (uint)rowIdx);
        return SetCellValue(excelSheet, excelCell, value, numberFormat);
    }

    /// <summary>
    /// Set a TimeOnly value in the cell.
    /// If the cell does not exist, it will be created.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="colIdx"></param>
    /// <param name="rowIdx"></param>
    /// <param name="value"></param>
    /// <param name="numberFormat"></param>
    /// <returns></returns>
    public bool SetCellValue(ExcelSheet excelSheet, int colIdx, int rowIdx, TimeOnly value, string numberFormat)
    {
        string colName = ExcelUtils.GetColumnName(colIdx);
        ExcelCell excelCell = CreateCell(excelSheet, colName, (uint)rowIdx);
        return SetCellValue(excelSheet, excelCell, value, numberFormat);
    }

    /// <summary>
    /// Set a double value in the cell.
    /// If the cell does not exist, it will be created.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="colIdx"></param>
    /// <param name="rowIdx"></param>
    /// <param name="value"></param>
    /// <param name="numberFormat"></param>
    /// <returns></returns>
    public bool SetCellValue(ExcelSheet excelSheet, int colIdx, int rowIdx, double value, string numberFormat)
    {
        string colName = ExcelUtils.GetColumnName(colIdx);
        ExcelCell excelCell = CreateCell(excelSheet, colName, (uint)rowIdx);
        return SetCellValue(excelSheet, excelCell, value, numberFormat);
    }

    #endregion Set cell value
}