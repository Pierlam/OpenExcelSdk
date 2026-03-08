using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenExcelSdk.Export;
using OpenExcelSdk.System;
using OpenExcelSdk.System.Export;
using System.Xml.Schema;

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
    StylesExtractor _stylesExtractor;


    public ExcelProcessor(): base()
    {
        _stylesExtractor = new StylesExtractor(this, _styleMgr);
    }

    /// <summary>
    /// Export all styles of the excel file into an output excel file.
    /// Style: CellFormat, Fill, Border, Font. 
    /// </summary>
    /// <param name="filenameIn"></param>
    /// <param name="filenameOut"></param>
    /// <returns></returns>
    public ExcelAllStylesExport ExportAllStyles(string filenameIn, string filenameOut)
    {
        ExcelFile excelFileIn = OpenExcelFile(filenameIn);

        // extract the styles from the input file
        ExcelAllStylesExport excelAllStyles = _stylesExtractor.Extract(excelFileIn);

        ExcelFile excelFileOut = CreateExcelFile(filenameOut, "Abstract");

        // tabpage 1: abstract
        ExcelAbstractExporter.Export(this, excelAllStyles, excelFileIn, excelFileOut);

        // tabpage 2: cells infos
        ExcelCellExporter.Export(this, excelAllStyles, excelFileIn, excelFileOut);

        // tabpage 3: SharedStrings
        ExcelSharedStringExporter.Export(this, excelAllStyles, excelFileOut);

        // tabpage 4: Styles
        ExcelStylesExporter.Export(this, excelAllStyles, excelFileOut);

        // tabpage 5: Fills
        ExcelFillsExporter.ExportFills(this, excelAllStyles, excelFileOut);

        // tabpage 6: Border
        ExcelBordersExporter.ExportBorders(this, excelAllStyles, excelFileOut);

        // tabpage 7: Font
        ExcelFontsExporter.ExportFonts(this, excelAllStyles, excelFileOut);


        CloseExcelFile(excelFileOut);

        return excelAllStyles;
    }

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
            // flushes and save modifications, release resources
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
    /// Return the number of sheets defined in the excel file.
    /// </summary>
    /// <param name="excelFile"></param>
    /// <param name="index"></param>
    /// <returns></returns>
    public int GetSheetCount(ExcelFile excelFile)
    {
        int? count= excelFile.WorkbookPart?.Workbook?.GetFirstChild<Sheets>()?.Elements<Sheet>()?.Count();
        if (count == null) return 0; 
        return count.Value;
    }

    /// <summary>
    /// Get the first sheet of the excel file.
    /// </summary>
    /// <param name="excelFile"></param>
    /// <param name="index"></param>
    /// <returns></returns>
    public ExcelSheet GetFirstSheet(ExcelFile excelFile)
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
    /// Get a row from the sheet  by index base1.
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
                // row does not exists, it's not an error
                return null;

            if (rowIndex < 1 || rowIndex > rows.Count())
                // row does not exists, it's not an error
                return null;

            Row row = rows.ElementAt(rowIndex-1);
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

    /// <summary>
    /// Return the number of cells of the row. If the row does not exists, return 0.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="rowIndex"></param>
    /// <returns></returns>
    public int GetRowCellsCount(ExcelSheet excelSheet, int rowIndex)
    {
        ExcelRow excelRow = GetRowAt(excelSheet, rowIndex);
        if (excelRow == null) return 0;
        return excelRow.Row.Elements<Cell>().Count();
    }

    /// <summary>
    /// Get cells of the row.
    /// </summary>
    /// <param name="excelRow"></param>
    /// <returns></returns>
    public List<ExcelCell> GetRowCells(ExcelSheet excelSheet, ExcelRow excelRow)
    {
        List<ExcelCell> listCell =new List<ExcelCell>();

        if(excelRow==null || excelRow.Row==null) return listCell;

        // Iterate through each cell in the row
        foreach (Cell cell in excelRow.Row.Elements<Cell>())
        {
            ExcelCell excelCell = new ExcelCell(excelSheet, cell);
            listCell.Add(excelCell);
        }
        return listCell;
    }

    /// <summary>
    /// Get cells of the row.
    /// </summary>
    /// <param name="excelRow"></param>
    /// <returns></returns>
    public List<ExcelCell> GetRowCells(ExcelSheet excelSheet, int rowIndex)
    {
        List<ExcelCell> listCell = new List<ExcelCell>();

        if(rowIndex < 1) return listCell;

        // Get the first worksheet
        SheetData sheetData = excelSheet.Worksheet.GetFirstChild<SheetData>();


        // Find the row by index
        Row row = sheetData.Elements<Row>()
                           .FirstOrDefault(r => r.RowIndex == rowIndex);
        if(row== null)  
            return listCell;

        // Iterate through each cell in the row
        foreach (Cell cell in row.Elements<Cell>())
        {
            ExcelCell excelCell = new ExcelCell(excelSheet, cell);
            listCell.Add(excelCell);
        }
        return listCell;
    }

    #endregion Get row

    #region Get Cell

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
        return GetCellAt(excelSheet, ExcelCellAddressUtils.ConvertAddress(colIdx, rowIdx));
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

        // check the cell address
        if (!ExcelCellAddressUtils.GetColumnAndRowIndex(cellReference, out _, out _))
            throw ExcelException.Create("GetCellAt", ExcelErrorCode.InvalidCellAddress, cellReference);

        try
        {
            Cell? cell = excelSheet.Worksheet?.Descendants<Cell>()?.Where(c => c.CellReference == cellReference).FirstOrDefault();
            if (cell == null)
                // no cell found, not an error
                return null;

            var excelCell = new ExcelCell(excelSheet, cell);

            // get the style of the cell
            excelCell.CellFormat = ExcelCellAddressUtils.GetCellFormat(excelSheet, excelCell);

            return excelCell;
        }
        catch (Exception ex)
        {
            throw ExcelException.Create("GetCellAt", ExcelErrorCode.UnableGetCell, cellReference, ex);
        }
    }

    /// <summary>
    /// Get cells count in the sheet.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <returns></returns>
    public int GetCellsCount(ExcelSheet excelSheet)
    {
        SheetData sheetData = excelSheet.Worksheet.Elements<SheetData>().FirstOrDefault();

        // Count all cells (including empty ones that exist in XML)
        return sheetData.Elements<Row>()
                                .SelectMany(r => r.Elements<Cell>())
                                .Count();
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

    #region Get something

    /// <summary>
    /// Get the foreground and background color of the cell.
    /// background and/or foreground color can be null if there is no color defined.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="excelCell"></param>
    /// <returns></returns>
    public ExcelCellColor GetCellColor(ExcelSheet excelSheet, string cellReference)
    {
        ExcelCell excelCell = GetCellAt(excelSheet, cellReference);
        if (excelCell == null) return null;

        return GetCellColor(excelSheet, excelCell);
    }

    /// <summary>
    /// Get the foreground and background color of the cell.
    /// background and/or foreground color can be null if there is no color defined.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="excelCell"></param>
    /// <returns></returns>
    public ExcelCellColor GetCellColor(ExcelSheet excelSheet, int colIdx, int rowIdx)
    {
        ExcelCell excelCell = GetCellAt(excelSheet, colIdx, rowIdx);
        if (excelCell == null) return null;

        return GetCellColor(excelSheet, excelCell);

    }

    #endregion

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
        // check the cell address
        if (!ExcelCellAddressUtils.GetColumnAndRowIndex(cellReference, out int colIdx, out int rowIdx))
            throw ExcelException.Create("CreateCell", ExcelErrorCode.InvalidCellAddress, cellReference);

        string colName = ExcelCellAddressUtils.GetColumnName(colIdx);

        return CreateCell(excelSheet, colName, (uint)rowIdx);
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
        string colName = ExcelCellAddressUtils.GetColumnName(colIdx);
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
    public bool RemoveCellAt(ExcelSheet excelSheet, string cellReference)
    {
        // check the cell address
        if (!ExcelCellAddressUtils.GetColumnAndRowIndex(cellReference, out int colIdx, out int rowIdx))
            throw ExcelException.Create("RemoveCellAt", ExcelErrorCode.InvalidCellAddress, cellReference);

        return RemoveCell(excelSheet, colIdx, rowIdx);
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
            throw ExcelException.Create("RemoveCell", ExcelErrorCode.UnableRemoveCell, ExcelCellAddressUtils.ConvertAddress(colIdx, rowIdx), ex);
        }
    }

    #endregion Remove cell

    #region Set cell value currency

    /// <summary>
    /// Sets the value of a specified cell in the given Excel sheet, applying the specified currency format and name.
    /// example: 123,45 €, currencyFormat= Currency, CurrencyName=Euro  
    /// </summary>
    /// <param name="excelSheet">The ExcelSheet object representing the sheet where the cell value will be set.</param>
    /// <param name="cellReference">The address of the cell to be updated, specified in standard Excel format (for example, "A1").</param>
    /// <param name="value">The numeric value to be set in the specified cell.</param>
    /// <param name="currencyFormat">The format to be applied to the cell value, defining how the currency will be displayed.</param>
    /// <param name="currencyName">The name of the currency to be used, which will be displayed alongside the value in the cell.</param>
    /// <returns>true if the cell value was successfully set; otherwise, false.</returns>
    public bool SetCellValueCurrency(ExcelSheet excelSheet, string cellReference, double value, CurrencyFormat currencyFormat, CurrencyName currencyName, int digitAfter)
    {
        // check the cell address
        if (!ExcelCellAddressUtils.GetColumnAndRowIndex(cellReference, out int colIdx, out int rowIdx))
            throw ExcelException.Create("SetCellValueEmpty", ExcelErrorCode.InvalidCellAddress, cellReference);

        return SetCellValueCurrency(excelSheet, colIdx, rowIdx, value, currencyFormat, currencyName,digitAfter);
    }

    /// <summary>
    /// Sets the value of a specified cell in the given Excel sheet, applying the specified currency format and name.
    /// example: 123,45 €, currencyFormat= Currency, CurrencyName=Euro  
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="colIdx"></param>
    /// <param name="rowIdx"></param>
    /// <param name="value"></param>
    /// <param name="digitAfter"></param>
    /// <param name="currencyFormat"></param>
    /// <param name="currencyName"></param>
    /// <returns></returns>
    public bool SetCellValueCurrency(ExcelSheet excelSheet, int colIdx, int rowIdx, double value, CurrencyFormat currencyFormat, CurrencyName currencyName, int digitAfter)
    {
        string colName = ExcelCellAddressUtils.GetColumnName(colIdx);
        // create the cell if it does not exist
        ExcelCell excelCell = CreateCell(excelSheet, colName, (uint)rowIdx);

        return SetCellValueCurrency(excelSheet, excelCell, value, currencyFormat, currencyName, digitAfter);

    }

    public bool SetCellValueCurrency(ExcelSheet excelSheet, ExcelCell excelCell, double value,  CurrencyFormat currencyFormat, CurrencyName currencyName, int digitAfter)
    {
        if(excelCell == null || excelCell.Cell == null)
        {
            // no cell at this address
            return false;
        }

        // format: Accounting -> exp: _-* #,##0.00\ "€"_-;\-* #,##0.00\ "€"_-;_-* "-"??\ "€"_-;_-@_-
        // format: Currency -> exp:  #,##0.00\ "€"
        if (!CurrencyMgr.CreateNumberFormat(currencyFormat, currencyName, digitAfter, out string numberFormat))
        { 
            // unable to create the currency
            return false;
        }
        return SetCellValue(excelSheet, excelCell, value, numberFormat);
    }

    #endregion

    #region Set cell value empty

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
        // check the cell address
        if (!ExcelCellAddressUtils.GetColumnAndRowIndex(cellReference, out int colIdx, out int rowIdx))
            throw ExcelException.Create("SetCellValueEmpty", ExcelErrorCode.InvalidCellAddress, cellReference);

        return SetCellValueEmpty(excelSheet, colIdx, rowIdx);
    }

    /// <summary>
    /// Empty/Clear a cell value.
    /// Keep the format: Alignement colors, border, ...
    /// If the cell contains a formula, remove it.
    /// It the cell is null, do nothing.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="cellReference"></param>
    /// <returns></returns>
    public bool SetCellValueEmpty(ExcelSheet excelSheet, int colIdx, int rowIdx)
    {
        ExcelCell excelCell = GetCellAt(excelSheet, colIdx,rowIdx);
        if (excelCell == null || excelCell.Cell == null)
        {
            // no cell at this address
            return false;
        }

        return SetCellValueEmpty(excelSheet, excelCell);
    }

    /// <summary>
    /// Empty/Clear a cell value.
    /// Keep the format: Alignement colors, border, ...
    /// If the cell contains a formula, remove it.
    /// It the cell is null, do nothing.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="cellReference"></param>
    /// <returns></returns>
    public bool SetCellValueEmpty(ExcelSheet excelSheet, ExcelCell excelCell)
    {
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

    #endregion

    #region Set cell value

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
        // check the cell address
        if (!ExcelCellAddressUtils.GetColumnAndRowIndex(cellReference, out int colIdx, out int rowIdx))
            throw ExcelException.Create("SetCellValue", ExcelErrorCode.InvalidCellAddress, cellReference);

        return SetCellValue(excelSheet, colIdx, rowIdx, value);
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
        // check the cell address
        if (!ExcelCellAddressUtils.GetColumnAndRowIndex(cellReference, out int colIdx, out int rowIdx)) 
            throw ExcelException.Create("SetCellValue", ExcelErrorCode.InvalidCellAddress, cellReference);

        return SetCellValue(excelSheet, colIdx, rowIdx, value);
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
        // check the cell address
        if (!ExcelCellAddressUtils.GetColumnAndRowIndex(cellReference, out int colIdx, out int rowIdx))
            throw ExcelException.Create("SetCellValue", ExcelErrorCode.InvalidCellAddress, cellReference);

        return SetCellValue(excelSheet, colIdx, rowIdx, value);
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
        // check the cell address
        if (!ExcelCellAddressUtils.GetColumnAndRowIndex(cellReference, out int colIdx, out int rowIdx))
            throw ExcelException.Create("SetCellValue", ExcelErrorCode.InvalidCellAddress, cellReference);

        return SetCellValue(excelSheet, colIdx, rowIdx, value, numberFormat);
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
        // check the cell address
        if (!ExcelCellAddressUtils.GetColumnAndRowIndex(cellReference, out int colIdx, out int rowIdx))
            throw ExcelException.Create("SetCellValue", ExcelErrorCode.InvalidCellAddress, cellReference);

        return SetCellValue(excelSheet, colIdx, rowIdx, value, numberFormat);
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
        // check the cell address
        if (!ExcelCellAddressUtils.GetColumnAndRowIndex(cellReference, out int colIdx, out int rowIdx))
            throw ExcelException.Create("SetCellValue", ExcelErrorCode.InvalidCellAddress, cellReference);

        return SetCellValue(excelSheet, colIdx, rowIdx, value, numberFormat);
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
        string colName = ExcelCellAddressUtils.GetColumnName(colIdx);
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
        string colName = ExcelCellAddressUtils.GetColumnName(colIdx);
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
        string colName = ExcelCellAddressUtils.GetColumnName(colIdx);
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
        string colName = ExcelCellAddressUtils.GetColumnName(colIdx);
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
        string colName = ExcelCellAddressUtils.GetColumnName(colIdx);
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
        string colName = ExcelCellAddressUtils.GetColumnName(colIdx);
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
        string colName = ExcelCellAddressUtils.GetColumnName(colIdx);
        ExcelCell excelCell = CreateCell(excelSheet, colName, (uint)rowIdx);
        return SetCellValue(excelSheet, excelCell, value, numberFormat);
    }

    #endregion Set cell value

    #region Copy Cell 

    /// <summary>
    ///   Copy the value of a cell to another cell from a source excel file/sheet to a destination excel file/sheet.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="cellReference"></param>
    /// <param name="excelFileDest"></param>
    /// <param name="cellReferenceDest"></param>
    /// <returns></returns>
    public bool CopyCellValue(ExcelSheet excelSheet, string cellReference, ExcelSheet excelSheetDest, string cellReferenceDest)
    {
        ExcelCell excelCell = GetCellAt(excelSheet, cellReference);

        // cell does not exist at source, so clear the destination cell 
        if (excelCell == null) 
        { 
            return SetCellValueEmpty(excelSheetDest, cellReferenceDest);
        }

        return CopyCellValue(excelSheet, excelCell, excelSheetDest, cellReferenceDest);
    }

    /// <summary>
    ///   Copy the value of a cell to another cell from a source excel file/sheet to a destination excel file/sheet.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="cellReference"></param>
    /// <param name="excelFileDest"></param>
    /// <param name="cellReferenceDest"></param>
    /// <returns></returns>
    public bool CopyCellValue(ExcelSheet excelSheet, ExcelCell excelCell, ExcelSheet excelSheetDest, string cellReferenceDest)
    {
        if (excelCell == null) return false;
        //if (excelCellDest == null) return false;

        string numberFormat = string.Empty;
        if (excelCell.CellFormat != null)
        {
            // is it a buit-in number format?
            _styleMgr.GetNumberFormat(excelSheet, (int)excelCell.CellFormat.NumberFormatId.Value, out numberFormat);
        }

        ExcelCellValue excelCellValue = GetCellValue(excelSheet, excelCell);

        ExcelCell excelCellDest = GetCellAt(excelSheetDest, cellReferenceDest);

        // source cell is empty, clear the destination cell
        if (excelCellValue.IsEmpty)
        {
            // destination cell is already null
            if (excelCellDest == null) return true;

            return SetCellValueEmpty(excelSheetDest, excelCellDest);
        }

        if (excelCellDest == null)
            excelCellDest = CreateCell(excelSheetDest, cellReferenceDest);

        if (excelCellValue.CellType == ExcelCellType.String)
        {
            return SetCellValue(excelSheetDest, excelCellDest, excelCellValue.StringValue);
        }

        if (excelCellValue.Currency != null)
        {
            // copy the currency too
            //SetCe
        }

        if (excelCellValue.CellType == ExcelCellType.Integer)
        {
            if (!string.IsNullOrEmpty(numberFormat))
                return SetCellValue(excelSheetDest, excelCellDest, excelCellValue.DoubleValue.Value, numberFormat);

            return SetCellValue(excelSheetDest, excelCellDest, excelCellValue.IntegerValue.Value);
        }

        if (excelCellValue.CellType == ExcelCellType.Double)
        {
            if (!string.IsNullOrEmpty(numberFormat))
                return SetCellValue(excelSheetDest, excelCellDest, excelCellValue.DoubleValue.Value, numberFormat);

            return SetCellValue(excelSheetDest, excelCellDest, excelCellValue.DoubleValue.Value);
        }

        if (excelCellValue.CellType == ExcelCellType.DateOnly)
        {
            return SetCellValue(excelSheetDest, excelCellDest, excelCellValue.DateOnlyValue.Value, numberFormat);
        }

        if (excelCellValue.CellType == ExcelCellType.DateTime)
        {
            return SetCellValue(excelSheetDest, excelCellDest, excelCellValue.DateTimeValue.Value, numberFormat);
        }

        if (excelCellValue.CellType == ExcelCellType.TimeOnly)
        {
            return SetCellValue(excelSheetDest, excelCellDest, excelCellValue.TimeOnlyValue.Value, numberFormat);
        }

        // cell value type not managed
        return false;
    }

    #endregion

    #region Set Cell something else

    /// <summary>
    /// Set a color to a cell.
    /// Technically set a color to the foreground property, null into background one and set Solid to the pattern property.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="excelCell"></param>
    /// <param name="rgb"></param>
    /// <returns></returns>
    public ExcelCellColor SetCellColor(ExcelSheet excelSheet,  string cellReference, string rgb)
    {
        ExcelCell excelCell = GetCellAt(excelSheet, cellReference);
        if (excelCell == null) return null;
        return SetCellColor(excelSheet, excelCell, rgb);
    }

    /// <summary>
    /// Set a color to a cell.
    /// Technically set a color to the foreground property, null into background one and set Solid to the pattern property.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="excelCell"></param>
    /// <param name="rgb"></param>
    /// <returns></returns>
    public ExcelCellColor SetCellColor(ExcelSheet excelSheet, int colIdx, int rowIdx, string rgb)
    {
        string colName = ExcelCellAddressUtils.GetColumnName(colIdx);
        ExcelCell excelCell = CreateCell(excelSheet, colName, (uint)rowIdx);
        return SetCellColor(excelSheet, excelCell, rgb);
    }

    #endregion

}