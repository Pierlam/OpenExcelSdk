using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Text = DocumentFormat.OpenXml.Spreadsheet.Text;

namespace OpenExcelSdk;

/// <summary>
/// Main class to process Excel files.
/// Open/close, create excel file.
/// Get/Create sheet. Get/Create cell.
/// Get/set cell value.
/// and more...
/// </summary>
public class ExcelProcessor: ExcelProcessorBase
{

    #region Open/Close Create Excel file

    /// <summary>
    /// Open an existing Excel file.
    /// </summary>
    /// <param name="fileName"></param>
    /// <param name="excelFile"></param>
    /// <param name="error"></param>
    /// <returns></returns>
    public bool Open(string fileName, out ExcelFile excelFile, out ExcelError error)
    {
        excelFile = null;
        error = null;

        if(string.IsNullOrWhiteSpace(fileName))
        {
            error = new ExcelError(ExcelErrorCode.FilenameNull);
            return false;
        }

        if (!File.Exists(fileName))
        {
            error = new ExcelError(ExcelErrorCode.FileNotFound);
            return false;
        }

        try
        {
            // Open the document for editing.
            SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, true);
            excelFile = new ExcelFile(fileName, document);
            return true;
        }
        catch (Exception ex)
        {
            error = new ExcelError(ExcelErrorCode.UnableOpenFile, ex);
            return false;
        }
    }

    /// <summary>
    /// Close an open excel file.
    /// </summary>
    /// <param name="excelFile"></param>
    /// <param name="error"></param>
    /// <returns></returns>
    public bool Close(ExcelFile excelFile, out ExcelError error)
    {
        error = null;
        try
        {
            excelFile.SpreadsheetDocument.Dispose();
            excelFile.SpreadsheetDocument = null;
            return true;
        }
        catch (Exception ex)
        {
            error = new ExcelError(ExcelErrorCode.UnableCloseFile, ex);
            return false;
        }
    }

    /// <summary>
    /// Create a new excel file with one sheet.
    /// The filename should not exists.
    /// exp: "C:\Files\MyExcel.xlsx"
    /// </summary>
    /// <param name="fileName"></param>
    /// <param name="excelFile"></param>
    /// <param name="error"></param>
    /// <returns></returns>
    public bool CreateExcelFile(string fileName, out ExcelFile excelFile, out ExcelError error)
    {
        return CreateExcelFile(fileName, Definitions.DefaultFirstSheetName, out excelFile, out error);
    }

    /// <summary>
    /// Create a new excel file with one sheet. Provide the sheet name.
    /// The filename should not exists.
    /// exp: "C:\Files\MyExcel.xlsx"
    ///
    /// https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/structure-of-a-spreadsheetml-document?tabs=cs
    /// </summary>
    /// <param name="fileName"></param>
    public bool CreateExcelFile(string fileName, string sheetName, out ExcelFile excelFile, out ExcelError error)
    {
        error = null;
        excelFile = null;

        if (File.Exists(fileName))
        {
            error = new ExcelError(ExcelErrorCode.FileAlreadyExists);
            return false;
        }

        if (string.IsNullOrWhiteSpace(sheetName))
        {
            error = new ExcelError(ExcelErrorCode.ValueNull);
            return false;
        }

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

            excelFile = new ExcelFile(fileName, spreadsheetDocument);

            return true;
        }
        catch (Exception ex)
        {
            error = new ExcelError(ExcelErrorCode.UnableCreateFile, ex);
            return false;
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
    /// <param name="excelSheet"></param>
    /// <param name="error"></param>
    /// <returns></returns>
    public bool CreateSheet(ExcelFile excelFile, string sheetName, out ExcelSheet excelSheet, out ExcelError error)
    {
        excelSheet = null;
        error = null;

        if (string.IsNullOrWhiteSpace(sheetName))
        {
            error = new ExcelError(ExcelErrorCode.UnableCreateSheet);
            return false;
        }

        if (excelFile.WorkbookPart.Workbook == null)
            // Add Sheets to the Workbook.
            excelFile.WorkbookPart.Workbook.AppendChild(new Sheets());

        var listSheets = excelFile.WorkbookPart.Workbook.Sheets.Elements<Sheet>();

        // Find the sheet with the matching name
        var sheet = listSheets.FirstOrDefault(s => s.Name != null && s.Name.Value.Equals(sheetName, StringComparison.InvariantCultureIgnoreCase));
        if (sheet != null)
        {
            // a sheet with the same already exists
            error = new ExcelError(ExcelErrorCode.UnableCreateSheet);
            return false;
        }

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

        excelSheet = new ExcelSheet(excelFile, sheet);
        excelSheet.Index = (int)sheetId;
        excelSheet.Name = sheet.Name;
        return true;
    }


    /// <summary>
    /// Get the first sheet of the excel file.
    /// </summary>
    /// <param name="excelFile"></param>
    /// <param name="index"></param>
    /// <param name="excelSheet"></param>
    /// <param name="error"></param>
    /// <returns></returns>
    public bool GetFirstSheet(ExcelFile excelFile, int index, out ExcelSheet excelSheet, out ExcelError error)
    {
        return GetSheetAt(excelFile, 0, out excelSheet, out error);
    }

    /// <summary>
    /// Get a sheet of the excel file by index base0.
    /// </summary>
    /// <param name="excelFile"></param>
    /// <param name="index"></param>
    /// <param name="excelSheet"></param>
    /// <param name="error"></param>
    /// <returns></returns>
    public bool GetSheetAt(ExcelFile excelFile, int index, out ExcelSheet excelSheet, out ExcelError error)
    {
        excelSheet = null;
        error = null;

        if (index < 0)
        {
            error = new ExcelError(ExcelErrorCode.IndexMustBePositive);
            return false;
        }

        if (excelFile == null)
        {
            error = new ExcelError(ExcelErrorCode.FilenameNull);
            return false;
        }

        try
        {
            Sheet? sheet = excelFile.WorkbookPart?.Workbook?.GetFirstChild<Sheets>()?.Elements<Sheet>()?.ElementAt<Sheet>(index);

            if (sheet == null)
            {
                error = new ExcelError(ExcelErrorCode.IndexWrong);
                return false;
            }

            excelSheet = new ExcelSheet(excelFile, sheet);
            excelSheet.Index = index;
            excelSheet.Name = sheet.Name;
            return true;
        }
        catch (Exception ex)
        {
            error = new ExcelError(ExcelErrorCode.UnableGetSheet, ex);
            excelSheet = null;
            return false;
        }
    }

    /// <summary>
    /// Get a sheet of the excel file by the name.
    /// </summary>
    /// <param name="excelFile"></param>
    /// <param name="index"></param>
    /// <param name="excelSheet"></param>
    /// <param name="error"></param>
    /// <returns></returns>
    public bool GetSheetByName(ExcelFile excelFile, string sheetName, out ExcelSheet excelSheet, out ExcelError error)
    {
        error = null;
        excelSheet = null;

        if (string.IsNullOrWhiteSpace(sheetName))
        {
            error = new ExcelError(ExcelErrorCode.ValueNull);
            return false;
        }

        try
        {
            // Get the Sheets collection
            var sheets = excelFile.WorkbookPart.Workbook.Sheets.Elements<Sheet>();

            // Find the sheet with the matching name
            var sheet = sheets.FirstOrDefault(s => s.Name != null && s.Name.Value.Equals(sheetName, StringComparison.OrdinalIgnoreCase));
            if (sheet == null)
                // no sheet with this name, not an error
                return false;

            excelSheet = new ExcelSheet(excelFile, sheet);
            return true;
        }
        catch (Exception ex)
        {
            error = new ExcelError(ExcelErrorCode.UnableGetSheet, ex);
            excelSheet = null;
            return false;
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
    /// <param name="excelRow"></param>
    /// <param name="error"></param>
    /// <returns></returns>
    public bool GetRowAt(ExcelSheet excelSheet, int rowIndex, out ExcelRow excelRow, out ExcelError error)
    {
        excelRow = null;
        error = null;

        try
        {
            var rows = excelSheet.Worksheet.Descendants<Row>();
            if (!rows.Any())
                // row doest not exists, it's not an error
                return true;

            if (rowIndex < 0 || rowIndex > rows.Count())
            {
                error = new ExcelError(ExcelErrorCode.IndexWrong);
                return false;
            }
            Row row = rows.ElementAt(rowIndex);
            excelRow = new ExcelRow(row);
            return true;
        }
        catch (Exception ex)
        {
            error = new ExcelError(ExcelErrorCode.UnableGetRow, ex);
            return false;
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

    #region Get infos

    /// <summary>
    /// Return the count of custom number formats in the excel sheet.
    /// It's style on cell value, exp: date, currency, percentage,...
    /// built-in number formats are not counted.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <returns></returns>
    public int GetCustomNumberFormatsCount(ExcelSheet excelSheet)
    {
        var stylesPart = excelSheet.ExcelFile.WorkbookPart.WorkbookStylesPart;
        if (stylesPart == null)
            return 0;
        if (stylesPart.Stylesheet == null)
            return 0;

        return stylesPart.Stylesheet.CellFormats.Elements().Count();
    }

    #endregion Get row 

    #region Get Cell at

    /// <summary>
    /// Get a cell in the sheet by col and row index, base1.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="colIdx"></param>
    /// <param name="rowIdx"></param>
    /// <param name="excelCell"></param>
    /// <param name="excelError"></param>
    /// <returns></returns>
    public bool GetCellAt(ExcelSheet excelSheet, int colIdx, int rowIdx, out ExcelCell excelCell, out ExcelError excelError)
    {
        // convert the col and the rox to an excel address
        return GetCellAt(excelSheet, ExcelUtils.ConvertAddress(colIdx, rowIdx), out excelCell, out excelError);
    }

    /// <summary>
    /// Get a cell in the sheet by the address name. exp: A1
    ///
    /// If the cell does not exists, return a  null ExcelCell without error.
    /// If the access to the cell fails, then return an error.
    /// https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/how-to-retrieve-the-values-of-cells-in-a-spreadsheet?tabs=cs-0%2Ccs-2%2Ccs-3%2Ccs-4%2Ccs-5%2Ccs-6%2Ccs-7%2Ccs-8%2Ccs-9%2Ccs-10%2Ccs-11%2Ccs
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="excelCell"></param>
    /// <param name="excelError"></param>
    /// <returns></returns>
    public bool GetCellAt(ExcelSheet excelSheet, string addressName, out ExcelCell excelCell, out ExcelError excelError)
    {
        excelCell = null;
        excelError = null;

        if (excelSheet == null)
        {
            excelError = new ExcelError(ExcelErrorCode.ObjectNull);
            return false;
        }

        try
        {
            Cell? cell = excelSheet.Worksheet?.Descendants<Cell>()?.Where(c => c.CellReference == addressName).FirstOrDefault();
            if (cell == null)
                // not an error
                return true;

            excelCell = new ExcelCell(excelSheet, cell);

            // get the style of the cell
            excelCell.CellFormat = GetCellFormat(excelSheet, excelCell);
            if (excelCell.Cell.StyleIndex != null)
            {
                var stylesPart = excelSheet.ExcelFile.WorkbookPart.WorkbookStylesPart;
                var cellFormat = (CellFormat)stylesPart.Stylesheet.CellFormats.ElementAt((int)excelCell.Cell.StyleIndex.Value);
            }

            return true;
        }
        catch (Exception ex)
        {
            excelError = new ExcelError(ExcelErrorCode.UnableGetCell, ex);
            return false;
        }
    }


    #endregion Get Cell at

    #region Get CellValue

    //public string GetCellValueAsString(ExcelSheet excelSheet, ExcelCell excelCell)
    //public string GetCellValueAsDouble(ExcelSheet excelSheet, ExcelCell excelCell)

    #endregion Get CellValue

    #region Get CellType

    /// <summary>
    /// Get the type of the cell value.
    /// If the cell is empty/blank, in some cases the type will be Undefined.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="addressName"></param>
    /// <returns></returns>
    public ExcelCellType GetCellType(ExcelSheet excelSheet, string addressName)
    {
        int colIdx= ExcelUtils.GetColumnIndex(addressName);
        int rowIdx = ExcelUtils.GetColumnIndex(addressName);

        return GetCellType(excelSheet, colIdx, rowIdx);
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
        bool res = GetCellAt(excelSheet, colIdx, rowIdx, out ExcelCell excelCell, out ExcelError excelError);
        if (!res) return ExcelCellType.Error;

        return GetCellType(excelSheet, excelCell);
    }

    /// <summary>
    /// Get the type, the value and the data format of the cell.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="excelCell"></param>
    /// <param name="excelCellValueMulti"></param>
    /// <param name="excelError"></param>
    /// <returns></returns>
    public bool GetCellTypeAndValue(ExcelSheet excelSheet, string addressName, out ExcelCell excelCell, out ExcelCellValueMulti excelCellValueMulti, out ExcelError excelError)
    {
        excelCell = null;
        excelCellValueMulti = null;
        excelError = null;
        bool res = GetCellAt(excelSheet, addressName, out excelCell, out excelError);
        if (!res) return false;

        return GetCellTypeAndValue(excelSheet, excelCell, out excelCellValueMulti, out excelError);
    }

    /// <summary>
    /// Geth the type, the value and the data format of cell.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="excelCell"></param>
    /// <param name="excelCellValueMulti"></param>
    /// <param name="excelError"></param>
    /// <returns></returns>
    public bool GetCellTypeAndValue(ExcelSheet excelSheet, int colIdx, int rowIdx, out ExcelCell excelCell, out ExcelCellValueMulti excelCellValueMulti, out ExcelError excelError)
    {
        excelCell = null;
        excelCellValueMulti = null;
        excelError = null;
        bool res = GetCellAt(excelSheet, colIdx, rowIdx, out excelCell, out excelError);
        if (!res) return false;

        return GetCellTypeAndValue(excelSheet, excelCell, out excelCellValueMulti, out excelError);
    }

    /// <summary>
    /// Get the style/CellFormat of the cell, if it has one.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="excelCell"></param>
    /// <returns></returns>
    public CellFormat GetCellFormat(ExcelSheet excelSheet, ExcelCell excelCell)
    {
        if (excelCell.Cell.StyleIndex == null)
            // no style, no cell format
            return null;

        var stylesPart = excelSheet.ExcelFile.WorkbookPart.WorkbookStylesPart;
        return (CellFormat)stylesPart.Stylesheet.CellFormats.ElementAt((int)excelCell.Cell.StyleIndex.Value);
    }

    #endregion Get CellType

    #region Create cell

    /// <summary>
    /// Given an address name e.g. A3, and a WorksheetPart, inserts a cell into the worksheet.
    /// If the cell already exists, returns it.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="addressName"></param>
    /// <param name="excelCell"></param>
    /// <param name="error"></param>
    /// <returns></returns>
    public bool CreateCell(ExcelSheet excelSheet, string addressName, out ExcelCell excelCell, out ExcelError error)
    {
        string colName = ExcelUtils.GetColumnName(ExcelUtils.GetColumnIndex(addressName));

        return CreateCell(excelSheet, colName, (uint)ExcelUtils.GetRowIndex(addressName), out excelCell, out error);
    }

    /// <summary>
    /// Given a column name, a row index, and a WorksheetPart, inserts a cell into the worksheet.
    /// If the cell already exists, returns it.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="colIdx"></param>
    /// <param name="rowIdx"></param>
    /// <param name="excelCell"></param>
    /// <param name="error"></param>
    /// <returns></returns>
    public bool CreateCell(ExcelSheet excelSheet, int colIdx, int rowIdx, out ExcelCell excelCell, out ExcelError error)
    {
        string colName = ExcelUtils.GetColumnName(colIdx);
        return CreateCell(excelSheet, colName, (uint)rowIdx, out excelCell, out error);
    }


    #endregion Create cell

    #region Remove cell

    /// <summary>
    /// Remove a cell in the sheet by col and row index, start at index 1.
    /// If there is not cell at the address, no eror will occur.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="colIdx"></param>
    /// <param name="rowIdx"></param>
    /// <param name="error"></param>
    /// <returns></returns>
    public bool RemoveCell(ExcelSheet excelSheet, string addressName, out ExcelError error)
    {
        return RemoveCell(excelSheet, ExcelUtils.GetColumnIndex(addressName), ExcelUtils.GetRowIndex(addressName), out error);
    }

    /// <summary>
    /// Remove a cell in the sheet by the address name. exp: A1
    /// If there is not cell at the address, no eror will occur.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="cellAddress"></param>
    /// <param name="error"></param>
    /// <returns></returns>
    public bool RemoveCell(ExcelSheet excelSheet, int colIdx, int rowIdx, out ExcelError error)
    {
        error = null;
        if (!GetCellAt(excelSheet, colIdx, rowIdx, out ExcelCell excelCell, out error))
            return false;
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
            error = new ExcelError(ExcelErrorCode.UnableRemoveCell, ex);
            return false;
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
    /// <param name="colIdx"></param>
    /// <param name="rowIdx"></param>
    /// <param name="error"></param>
    /// <returns></returns>
    public bool SetCellValueEmpty(ExcelSheet excelSheet, string addressName, out ExcelError error)
    {
        if (!GetCellAt(excelSheet, addressName, out ExcelCell excelCell, out error))
            return false;

        if (excelCell == null || excelCell.Cell == null) return true;

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
    /// <param name="error"></param>
    /// <returns></returns>
    public bool SetCellValue(ExcelSheet excelSheet, string addressName, string value, out ExcelError error)
    {
        return SetCellValue(excelSheet, ExcelUtils.GetColumnIndex(addressName), ExcelUtils.GetRowIndex(addressName), value, out error);
    }

    /// <summary>
    /// Set an int value in the cell.
    /// If the cell does not exist, it will be created.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="colIdx"></param>
    /// <param name="rowIdx"></param>
    /// <param name="value"></param>
    /// <param name="error"></param>
    /// <returns></returns>
    public bool SetCellValue(ExcelSheet excelSheet, string addressName, int value, out ExcelError error)
    {
        return SetCellValue(excelSheet, ExcelUtils.GetColumnIndex(addressName), ExcelUtils.GetRowIndex(addressName), value, out error);
    }

    /// <summary>
    /// Set a double value in the cell.
    /// If the cell does not exist, it will be created.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="colIdx"></param>
    /// <param name="rowIdx"></param>
    /// <param name="value"></param>
    /// <param name="error"></param>
    /// <returns></returns>
    public bool SetCellValue(ExcelSheet excelSheet, string addressName, double value, out ExcelError error)
    {
        return SetCellValue(excelSheet, ExcelUtils.GetColumnIndex(addressName), ExcelUtils.GetRowIndex(addressName), value, out error);
    }

    /// <summary>
    /// Set a DateOnly value in the cell.
    /// If the cell does not exist, it will be created.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="colIdx"></param>
    /// <param name="rowIdx"></param>
    /// <param name="value"></param>
    /// <param name="error"></param>
    /// <returns></returns>
    public bool SetCellValue(ExcelSheet excelSheet, string addressName, DateTime value, string format, out ExcelError error)
    {
        return SetCellValue(excelSheet, ExcelUtils.GetColumnIndex(addressName), ExcelUtils.GetRowIndex(addressName), value, format, out error);
    }

    /// <summary>
    /// Set a TimeOnly value in the cell.
    /// If the cell does not exist, it will be created.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="colIdx"></param>
    /// <param name="rowIdx"></param>
    /// <param name="value"></param>
    /// <param name="error"></param>
    /// <returns></returns>
    public bool SetCellValue(ExcelSheet excelSheet, string addressName, TimeOnly value, string format, out ExcelError error)
    {
        return SetCellValue(excelSheet, ExcelUtils.GetColumnIndex(addressName), ExcelUtils.GetRowIndex(addressName), value, format, out error);
    }

    /// <summary>
    /// Set a DateOnly value in the cell.
    /// If the cell does not exist, it will be created.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="colIdx"></param>
    /// <param name="rowIdx"></param>
    /// <param name="value"></param>
    /// <param name="error"></param>
    /// <returns></returns>
    public bool SetCellValue(ExcelSheet excelSheet, string addressName, DateOnly value, string format, out ExcelError error)
    {
        return SetCellValue(excelSheet, ExcelUtils.GetColumnIndex(addressName), ExcelUtils.GetRowIndex(addressName), value, format, out error);
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
    /// <param name="error"></param>
    /// <returns></returns>
    public bool SetCellValueEmpty(ExcelSheet excelSheet, int colIdx, int rowIdx, out ExcelError error)
    {
        if (!GetCellAt(excelSheet, colIdx, rowIdx, out ExcelCell excelCell, out error))
            return false;

        if (excelCell == null || excelCell.Cell == null) return true;

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
    /// <param name="error"></param>
    /// <returns></returns>
    public bool SetCellValue(ExcelSheet excelSheet, int colIdx, int rowIdx, string value, out ExcelError error)
    {
        string colName = ExcelUtils.GetColumnName(colIdx);
        if (!CreateCell(excelSheet, colName, (uint)rowIdx, out ExcelCell excelCell, out error))
            return false;
        return SetCellValue(excelSheet, excelCell, value, out error);
    }

    /// <summary>
    /// Set an int value in the cell.
    /// If the cell does not exist, it will be created.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="colIdx"></param>
    /// <param name="rowIdx"></param>
    /// <param name="value"></param>
    /// <param name="error"></param>
    /// <returns></returns>
    public bool SetCellValue(ExcelSheet excelSheet, int colIdx, int rowIdx, int value, out ExcelError error)
    {
        string colName = ExcelUtils.GetColumnName(colIdx);
        if (!CreateCell(excelSheet, colName, (uint)rowIdx, out ExcelCell excelCell, out error))
            return false;
        return SetCellValue(excelSheet, excelCell, value, out error);
    }

    /// <summary>
    /// Set a double value in the cell.
    /// If the cell does not exist, it will be created.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="colIdx"></param>
    /// <param name="rowIdx"></param>
    /// <param name="value"></param>
    /// <param name="error"></param>
    /// <returns></returns>
    public bool SetCellValue(ExcelSheet excelSheet, int colIdx, int rowIdx, double value, out ExcelError error)
    {
        string colName = ExcelUtils.GetColumnName(colIdx);
        if (!CreateCell(excelSheet, colName, (uint)rowIdx, out ExcelCell excelCell, out error))
            return false;
        return SetCellValue(excelSheet, excelCell, value, out error);
    }

    /// <summary>
    /// Set a DateOnly value in the cell.
    /// If the cell does not exist, it will be created.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="colIdx"></param>
    /// <param name="rowIdx"></param>
    /// <param name="value"></param>
    /// <param name="error"></param>
    /// <returns></returns>
    public bool SetCellValue(ExcelSheet excelSheet, int colIdx, int rowIdx, DateOnly value, string format, out ExcelError error)
    {
        string colName = ExcelUtils.GetColumnName(colIdx);
        if (!CreateCell(excelSheet, colName, (uint)rowIdx, out ExcelCell excelCell, out error))
            return false;
        return SetCellValue(excelSheet, excelCell, value, format, out error);
    }

    /// <summary>
    /// Set a DateTime value in the cell.
    /// If the cell does not exist, it will be created.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="colIdx"></param>
    /// <param name="rowIdx"></param>
    /// <param name="value"></param>
    /// <param name="error"></param>
    /// <returns></returns>
    public bool SetCellValue(ExcelSheet excelSheet, int colIdx, int rowIdx, DateTime value, string format, out ExcelError error)
    {
        string colName = ExcelUtils.GetColumnName(colIdx);
        if (!CreateCell(excelSheet, colName, (uint)rowIdx, out ExcelCell excelCell, out error))
            return false;
        return SetCellValue(excelSheet, excelCell, value, format, out error);
    }

    /// <summary>
    /// Set a TimeOnly value in the cell.
    /// If the cell does not exist, it will be created.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="colIdx"></param>
    /// <param name="rowIdx"></param>
    /// <param name="value"></param>
    /// <param name="error"></param>
    /// <returns></returns>
    public bool SetCellValue(ExcelSheet excelSheet, int colIdx, int rowIdx, TimeOnly value, string format, out ExcelError error)
    {
        string colName = ExcelUtils.GetColumnName(colIdx);
        if (!CreateCell(excelSheet, colName, (uint)rowIdx, out ExcelCell excelCell, out error))
            return false;
        return SetCellValue(excelSheet, excelCell, value, format, out error);
    }

    /// <summary>
    /// Set a double value in the cell.
    /// If the cell does not exist, it will be created.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="colIdx"></param>
    /// <param name="rowIdx"></param>
    /// <param name="value"></param>
    /// <param name="error"></param>
    /// <returns></returns>
    public bool SetCellValue(ExcelSheet excelSheet, int colIdx, int rowIdx, double value, string format, out ExcelError error)
    {
        string colName = ExcelUtils.GetColumnName(colIdx);
        if (!CreateCell(excelSheet, colName, (uint)rowIdx, out ExcelCell excelCell, out error))
            return false;
        return SetCellValue(excelSheet, excelCell, value, format, out error);
    }

    #endregion Set cell values
}