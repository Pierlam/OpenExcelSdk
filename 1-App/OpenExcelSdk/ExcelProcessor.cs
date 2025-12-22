using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.VariantTypes;
using OpenExcelSdk.System;
using System;
using static System.Runtime.InteropServices.JavaScript.JSType;
using NumberingFormat = DocumentFormat.OpenXml.Spreadsheet.NumberingFormat;
using Text = DocumentFormat.OpenXml.Spreadsheet.Text;

namespace OpenExcelSdk;

/// <summary>
/// Main class to process Excel files.
/// </summary>
public class ExcelProcessor
{
    StyleMgr _styleMgr = new StyleMgr();

    #region Open/Close Create Excel file
    public bool Open(string fileName, out ExcelFile excelFile, out ExcelError error)
    {
        excelFile = null;
        error = null;

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

    public bool Close(ExcelFile excelFile, out ExcelError error)
    {
        error = null;
        try
        {
            // TODO: add try-catch
            excelFile.SpreadsheetDocument.Dispose();
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
        return CreateExcelFile(fileName, "Sheet", out excelFile, out error);
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
    #endregion

    #region get sheet, row ,cell

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
    /// Get the sheet of the excel file by index base0.
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
            error = new ExcelError(ExcelErrorCode.FileNull);
            return false;
        }

        Sheet? sheet = excelFile.WorkbookPart?.Workbook?.GetFirstChild<Sheets>()?.Elements<Sheet>()?.ElementAt<Sheet>(index);

        if (sheet == null)
        {
            error = new ExcelError(ExcelErrorCode.IndexWrong);
            return false;
        }

        excelSheet = new ExcelSheet(excelFile, sheet);
        return true;
    }

    //public bool GetSheetByName(ExcelFileOXml excelFile, string sheetName index, out ExcelSheetOXml excelSheetOXml, out OXmlError error)
    //IEnumerable<Sheet>? sheets = excelFile.WorkbookPart?.Workbook?.GetFirstChild<Sheets>()?.Elements<Sheet>()?.Where(s => s.Name is not null && s.Name == sheetName);

    /// <summary>
    /// Get a row from the sheet  by index base0.
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
        //WorksheetPart worksheetPart = (WorksheetPart)excelSheet.ExcelFile.WorkbookPart.GetPartById(excelSheet.Sheet.Id);

        var rows = excelSheet.Worksheet.Descendants<Row>();
        if (!rows.Any()) return false;

        if (rowIndex < 0 || rowIndex > rows.Count())
        {
            error = new ExcelError(ExcelErrorCode.IndexWrong);
            return false;
        }
        Row row = rows.ElementAt(rowIndex);
        excelRow = new ExcelRow(row);
        return true;
    }

    //public IExcelCell GetCellAt(IExcelSheet excelSheet, int rowNum, int colNum)

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

        if(excelSheet == null)
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
            if (excelCell.Cell.StyleIndex!=null)
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

    /// <summary>
    /// Get the value of the cell as a string.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="cell"></param>
    /// <param name="stringValue"></param>
    /// <param name="error"></param>
    /// <returns></returns>
    public string GetCellValueAsString(ExcelSheet excelSheet, ExcelCell excelCell)
    {
        bool res = GetCellTypeAndValue(excelSheet, excelCell, out ExcelCellValueMulti excelCellValueMulti, out ExcelError excelError);
        if(!res)
            return string.Empty;

        if (excelCellValueMulti.CellType== ExcelCellType.String) 
            return excelCellValueMulti.StringValue;

        if (excelCellValueMulti.CellType == ExcelCellType.Integer)
            return excelCellValueMulti.IntegerValue.ToString();

        if (excelCellValueMulti.CellType == ExcelCellType.Double)
            return excelCellValueMulti.DoubleValue.ToString();

        if (excelCellValueMulti.CellType == ExcelCellType.DateTime)
            return excelCellValueMulti.DateTimeValue.ToString();

        if (excelCellValueMulti.CellType == ExcelCellType.DateOnly)
            return excelCellValueMulti.DateOnlyValue.ToString();

        if (excelCellValueMulti.CellType == ExcelCellType.TimeOnly)
            return excelCellValueMulti.TimeOnlyValue.ToString();
        return string.Empty;
    }

    public double GetCellValueAsDouble(ExcelSheet excelSheet, ExcelCell excelCell)
    {
        bool res = GetCellTypeAndValue(excelSheet, excelCell, out ExcelCellValueMulti excelCellValueMulti, out ExcelError excelError);
        if (!res)
            return 0;

        if (excelCellValueMulti.CellType == ExcelCellType.String)
            return 0;

        if (excelCellValueMulti.CellType == ExcelCellType.Integer)
            return excelCellValueMulti.IntegerValue.Value;

        if (excelCellValueMulti.CellType == ExcelCellType.Double)
            return excelCellValueMulti.DoubleValue.Value;

        if (excelCellValueMulti.CellType == ExcelCellType.DateTime)
            return excelCellValueMulti.DateTimeValue.Value.ToOADate();

        if (excelCellValueMulti.CellType == ExcelCellType.DateOnly)
            // not possible, so return zero
            return 0;

        if (excelCellValueMulti.CellType == ExcelCellType.TimeOnly)
            // not possible, so return zero
            return 0;

        return 0;
    }

    /// <summary>
    /// Get the type of cell value.
    /// GetCellTypeOfValue
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="excelCell"></param>
    /// <returns></returns>
    public ExcelCellType GetCellType(ExcelSheet excelSheet, ExcelCell excelCell)
    {
        bool res= GetCellTypeAndValue(excelSheet, excelCell, out ExcelCellValueMulti excelCellValueMulti, out ExcelError excelError);
        if(res)return excelCellValueMulti.CellType;
        return ExcelCellType.Error;
    }

    public bool GetCellTypeAndValue(ExcelSheet excelSheet, int colIdx, int rowIdx, out ExcelCell excelCell, out ExcelCellValueMulti excelCellValueMulti, out ExcelError excelError)
    { 
        excelCell = null;
        excelCellValueMulti = null;
        excelError = null;
        bool res = GetCellAt(excelSheet, colIdx, rowIdx, out excelCell, out excelError);
        if (!res) return false;

        return GetCellTypeAndValue(excelSheet, excelCell, out  excelCellValueMulti, out excelError);
    }

    /// <summary>
    /// Geth the type, the value and the data format of cell.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="excelCell"></param>
    /// <param name="excelCellValueMulti"></param>
    /// <param name="excelError"></param>
    /// <returns></returns>
    public bool GetCellTypeAndValue(ExcelSheet excelSheet, ExcelCell excelCell, out ExcelCellValueMulti excelCellValueMulti, out ExcelError excelError)
    {
        excelCellValueMulti = null;
        excelError = null;

        if (excelSheet == null || excelCell == null)
        {
            excelError = new ExcelError(ExcelErrorCode.ObjectNull);
            return false;
        }

        // no cell, is null, not an error
        if (excelCell.Cell == null)
        {
            excelCellValueMulti= new ExcelCellValueMulti();
            return true;
        }

        bool isTheCase;

        //--cell datatype is defined?
        if (!GetCellStringValue(excelSheet, excelCell, out isTheCase, out excelCellValueMulti, out excelError))
            return false;
        if (isTheCase) return true;

        // get the number format id
        if (!_styleMgr.GetCellNumberFormatId(excelSheet, excelCell, out uint numFmtId))
        {
            excelError = new ExcelError(ExcelErrorCode.TypeWrong);
            return false;
        }

        string value = excelCell.Cell.InnerText;
        double valDouble;
        int valInt;

        // is it a built-in format?
        if (BuiltInNumberFormatMgr.GetFormatAndType(numFmtId, out string dataFormat, out ExcelCellType cellType))
        {
            if (cellType == ExcelCellType.Integer)
                return CreateValueInteger(value, dataFormat, out excelCellValueMulti, out excelError);

            if (cellType == ExcelCellType.Double)
                return CreateValueDouble(value, dataFormat, out excelCellValueMulti, out excelError);

            if (cellType == ExcelCellType.DateOnly)
                return CreateValueDateOnly(value, dataFormat, out excelCellValueMulti, out excelError);

            if (cellType == ExcelCellType.DateTime)
                return CreateValueDateTime(value, dataFormat, out excelCellValueMulti, out excelError);

            if (cellType == ExcelCellType.TimeOnly)
                return CreateValueTimeOnly(value, dataFormat, out excelCellValueMulti, out excelError);

            excelError = new ExcelError(ExcelErrorCode.TypeWrong);
            return false;
        }

        // Try to get custom format if exists
        if (_styleMgr.GetCustomNumberFormat(excelSheet, numFmtId, out dataFormat))
        {
            // then determine the type from the data format: date, number,...
            cellType = GetCellType(dataFormat);

            if(cellType == ExcelCellType.DateTime)
                return CreateValueDateTime(value, dataFormat, out excelCellValueMulti, out excelError);

            if (cellType == ExcelCellType.DateOnly)
                return CreateValueDateOnly(value, dataFormat, out excelCellValueMulti, out excelError);

            if (cellType == ExcelCellType.TimeOnly)
                return CreateValueTimeOnly(value, dataFormat, out excelCellValueMulti, out excelError);

            if (cellType == ExcelCellType.Double)
                return CreateValueDouble(value, dataFormat, out excelCellValueMulti, out excelError);

            excelError = new ExcelError(ExcelErrorCode.TypeWrong);
            return false;
        }

        // on value in the cell?
        string cellValue = excelCell.Cell.InnerText;
        if (cellValue == string.Empty)
        {
            excelCellValueMulti = new ExcelCellValueMulti();
            return true;
        }

        // is it an int?
        bool res = int.TryParse(cellValue, out valInt);
        if (res)
        {
            excelCellValueMulti = new ExcelCellValueMulti(valInt);
            isTheCase = true;
            return true;
        }

        // is it a double?  cultureInfo prb: replace . by ,
        cellValue = cellValue.Replace('.', ',');
        res = double.TryParse(cellValue, out valDouble);
        if (res)
        {
            excelCellValueMulti = new ExcelCellValueMulti(valDouble);
            isTheCase = true;
            return true;
        }

        // not able to find the type
        excelError = new ExcelError(ExcelErrorCode.TypeWrong);
        return false;
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

    #endregion


    #region Create sheet, row ,cell

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

    /// <summary>
    /// Given a column name, a row index, and a WorksheetPart, inserts a cell into the worksheet. 
    /// If the cell already exists, returns it. 
    /// Exp: "A", 1 for cell A1.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="columnName"></param>
    /// <param name="rowIndex"></param>
    /// <param name="excelCell"></param>
    /// <param name="error"></param>
    /// <returns></returns>
    public bool CreateCell(ExcelSheet excelSheet, string columnName, uint rowIndex, out ExcelCell excelCell, out ExcelError error)
    {
        excelCell = null;
        error = null;

        Worksheet worksheet = excelSheet.Worksheet;
        SheetData? sheetData = worksheet.GetFirstChild<SheetData>();
        string cellReference = columnName + rowIndex;

        // If the worksheet does not contain a row with the specified row index, insert one.
        Row row;

        if (sheetData?.Elements<Row>().Where(r => r.RowIndex is not null && r.RowIndex == rowIndex).Count() != 0)
        {
            row = sheetData!.Elements<Row>().Where(r => r.RowIndex is not null && r.RowIndex == rowIndex).First();
        }
        else
        {
            row = new Row() { RowIndex = rowIndex };
            sheetData.Append(row);
        }

        // If there is not a cell with the specified column name, insert one.  
        if (row.Elements<Cell>().Where(c => c.CellReference is not null && c.CellReference.Value == columnName + rowIndex).Count() > 0)
        {
            Cell cell = row.Elements<Cell>().Where(c => c.CellReference is not null && c.CellReference.Value == cellReference).First();
            excelCell = new ExcelCell(excelSheet, cell);
            return true;
        }
        else
        {
            // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
            Cell? refCell = null;

            foreach (Cell cell in row.Elements<Cell>())
            {
                if (string.Compare(cell.CellReference?.Value, cellReference, true) > 0)
                {
                    refCell = cell;
                    break;
                }
            }

            Cell newCell = new Cell() { CellReference = cellReference };
            row.InsertBefore(newCell, refCell);

            excelCell = new ExcelCell(excelSheet, newCell);
            return true;
        }
    }


    #endregion

    #region Set cell value

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
        if(!CreateCell(excelSheet, colName, (uint)rowIdx, out ExcelCell excelCell, out error))
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

    /// <summary>
    /// Set a string value in the existing cell.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="excelCell"></param>
    /// <param name="value"></param>
    /// <param name="error"></param>
    /// <returns></returns>
    public bool SetCellValue(ExcelSheet excelSheet, ExcelCell excelCell, string value, out ExcelError error)
    {
        error = null;

        if(excelCell == null || excelCell.Cell == null)
        {
            error = new ExcelError(ExcelErrorCode.ObjectNull);
            return false;
        }

        try
        {
            WorkbookPart workbookPart = excelSheet.ExcelFile.WorkbookPart;

            // get the table
            SharedStringTablePart shareStringPart = GetOrCreateSharedStringTablePart(excelSheet.ExcelFile.WorkbookPart);

            // Insert the text into the SharedStringTablePart
            int index = InsertSharedStringItem(value, shareStringPart);

            // Set the value of cell A1
            excelCell.Cell.CellValue = new CellValue(index.ToString());
            excelCell.Cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);

            // remove formula if it's there
            _styleMgr.RemoveFormula(excelSheet, excelCell);

            // no cell format, nothing more to do
            if (!_styleMgr.HasCellFormat(excelSheet, excelCell))return true;

            // all other style than format (no border, no color,...) are null, clear the style of the cell
            if (_styleMgr.AllOthersStyleThanFormatAreNull(excelSheet, excelCell))
            {
                excelCell.Cell.StyleIndex = 0;
                return true;
            }

            // duplicate the style to update the CellFormat
            _styleMgr.UpdateCellStyleNumberFormatId(excelSheet, excelCell, null);

            return true;
        }
        catch (Exception ex)
        {
            error = new ExcelError(ExcelErrorCode.UnableSetCellValue, ex);
            return false;
        }
    }

    //public bool SetCellValue(ExcelSheet excelSheet, ExcelCell excelCell, int value, out ExcelError error)
    //{
    //    error = null;
    //    // Important: store as number
    //    excelCell.Cell.DataType = CellValues.Number;
    //    // Must be string in XML
    //    excelCell.Cell.CellValue = new CellValue(value.ToString());

    //    // remove formula if it's there
    //    _styleMgr.RemoveFormula(excelSheet, excelCell);

    //    // no cell format, nothing more to do
    //    if (!_styleMgr.HasCellFormat(excelSheet, excelCell)) return true;

    //    // all other style than format (no border, no color,...) are null, clear the style of the cell
    //    if (_styleMgr.AllOthersStyleThanFormatAreNull(excelSheet, excelCell))
    //    {
    //        excelCell.Cell.StyleIndex = 0;
    //        return true;
    //    }

    //    // duplicate the style to update the CellFormat
    //    _styleMgr.UpdateCellStyleNumberFormatId(excelSheet, excelCell, null);

    //    return true;
    //}

    public bool SetCellValue(ExcelSheet excelSheet, ExcelCell excelCell, int value, string format, out ExcelError error)
    {
        error = null;

        uint formatId;

        // get (built-in or custom) or create the format (custom)
        if (!_styleMgr.GetOrCreateNumberFormat(excelSheet, format, out formatId, out error))
            return false;

        return SetCellValueAndNumberFormatId(excelSheet, excelCell, value, formatId, out error);
    }


    public bool SetCellValue(ExcelSheet excelSheet, ExcelCell excelCell, double value, string format, out ExcelError error)
    {
        error = null;

        uint formatId;

        // get (built-in or custom) or create the format (custom)
        if(!_styleMgr.GetOrCreateNumberFormat(excelSheet, format, out formatId, out error))
            return false;

        return SetCellValueAndNumberFormatId(excelSheet, excelCell, value, formatId, out error);
    }

    public bool SetCellValue(ExcelSheet excelSheet, ExcelCell excelCell, int value, out ExcelError error)
    {
        return SetCellValueAndNumberFormatId(excelSheet, excelCell, value, null, out error);
    }

    public bool SetCellValue(ExcelSheet excelSheet, ExcelCell excelCell, double value, out ExcelError error)
    {
        return SetCellValueAndNumberFormatId(excelSheet, excelCell, value, null, out error);
    }

    /// <summary>
    /// Set cell value as double.
    /// Keep some aprt of the style: border, color, font...
    /// but clear the number format -> style/CellFormat/NumberingFormat
    /// TODO: to remove or rework.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="excelCell"></param>
    /// <param name="value"></param>
    /// <param name="error"></param>
    /// <returns></returns>
    public bool SetCellValueAndNumberFormatId(ExcelSheet excelSheet, ExcelCell excelCell, double value, uint? numberFormatId, out ExcelError error)
    {
        error = null;
        // Important: store as number
        excelCell.Cell.DataType = CellValues.Number;
        excelCell.Cell.CellValue = new CellValue(value.ToString(global::System.Globalization.CultureInfo.InvariantCulture));

        // remove formula if it's there
        _styleMgr.RemoveFormula(excelSheet, excelCell);

        // TODO: numberFormatId

        // no cell format, nothing more to do
        if (!_styleMgr.HasCellFormat(excelSheet, excelCell)) return true;

        // the cell contains the expected number format 
        _styleMgr.GetCellNumberFormatId(excelSheet, excelCell, out uint numberFormatIdCell);
        if(numberFormatIdCell == (numberFormatId ?? 0)) return true;

        // all other style than format (no border, no color,...) are null, clear the style of the cell
        if (numberFormatId ==null && _styleMgr.AllOthersStyleThanFormatAreNull(excelSheet, excelCell))
        {
            // TODO:
            excelCell.Cell.StyleIndex = 0;
            return true;
        }

        // duplicate the style to update the CellFormat
        _styleMgr.UpdateCellStyleNumberFormatId(excelSheet, excelCell, numberFormatId);

        return true;
    }

    #endregion

    #region privates methods

    bool GetCellStringValue(ExcelSheet excelSheet, ExcelCell excelCell, out bool isTheCase, out ExcelCellValueMulti excelCellValueMulti, out ExcelError error)
    {
        isTheCase = false;
        excelCellValueMulti = null;
        error = null;
        string cellValue;

        if (excelCell.Cell.DataType == null) return true;

        if (excelCell.Cell.DataType.Value == CellValues.SharedString)
        {
            if (!GetSharedStringValue(excelSheet, excelCell, out cellValue))
            {
                error = new ExcelError(ExcelErrorCode.UnableGetCellStringValue);
                return false;
            }
            excelCellValueMulti = new ExcelCellValueMulti(cellValue);
            isTheCase = true;
            return true;
        }

        if (excelCell.Cell.DataType.Value == CellValues.InlineString)
        {
            cellValue = excelCell.Cell.InlineString?.Text?.Text ?? string.Empty;
            if (cellValue == null)
            {
                error = new ExcelError(ExcelErrorCode.UnableGetCellStringValue);
                return false;
            }
            excelCellValueMulti = new ExcelCellValueMulti(cellValue);
            isTheCase = true;
            return true;
        }

        if (excelCell.Cell.DataType.Value == CellValues.String)
        {
            string value = excelCell.Cell.InnerText;
            if (value == null) value = string.Empty;
            excelCellValueMulti = new ExcelCellValueMulti(value);
            isTheCase = true;
            return true;
        }

        // not a string, bye
        return true;
    }

    bool GetSharedStringValue(ExcelSheet excelSheet, ExcelCell excelCell, out string stringValue)
    {
        stringValue = string.Empty;

        if (excelCell.Cell.DataType.Value != CellValues.SharedString) return false;

        // For shared strings, look up the value in the shared strings table.
        var stringTable = excelSheet.ExcelFile.WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

        // If the shared string table is missing, something is wrong. 
        // Return the index that is in the cell. 
        // Otherwise, look up the correct text in the table.
        if (stringTable is not null)
        {
            string value = excelCell.Cell.CellValue.InnerText;
            stringValue = stringTable.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
            return true;
        }
        return false;
    }

    /// <summary>
    /// Get the SharedStringTablePart. If it does not exist, create a new one.
    /// </summary>
    /// <param name="excelFile"></param>
    /// <returns></returns>
    private SharedStringTablePart GetOrCreateSharedStringTablePart(WorkbookPart workbookPart)
    {
        if (workbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
            return workbookPart.GetPartsOfType<SharedStringTablePart>().First();

        return workbookPart.AddNewPart<SharedStringTablePart>();
    }

    /// <summary>
    /// Given text and a SharedStringTablePart, creates a SharedStringItem with the specified text 
    ///  and inserts it into the SharedStringTablePart. If the item already exists, returns its index.
    /// </summary>
    /// <param name="text"></param>
    /// <param name="shareStringPart"></param>
    /// <returns></returns>
    private int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
    {
        // If the part does not contain a SharedStringTable, create one.
        shareStringPart.SharedStringTable ??= new SharedStringTable();

        int i = 0;

        // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
        foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
        {
            if (item.InnerText == text)
                return i;

            i++;
        }

        // The text does not exist in the part. Create the SharedStringItem and return its index.
        shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new Text(text)));

        return i;
    }

    static bool CreateValueInteger(string value, string dataFormat, out ExcelCellValueMulti excelCellValueMulti, out ExcelError excelError)
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

    static bool CreateValueDouble(string value, string dataFormat, out ExcelCellValueMulti excelCellValueMulti, out ExcelError excelError)
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

    static bool CreateValueDateOnly(string value, string dataFormat, out ExcelCellValueMulti excelCellValueMulti, out ExcelError excelError)
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

    static bool CreateValueDateTime(string value, string dataFormat, out ExcelCellValueMulti excelCellValueMulti, out ExcelError excelError)
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

    static bool CreateValueTimeOnly(string value, string dataFormat, out ExcelCellValueMulti excelCellValueMulti, out ExcelError excelError)
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

    /// <summary>
    /// Get the type of cell from the data format.
    /// exp:
    /// "dd/mm/yyyy\\ hh:mm:ss" , it's a DateTime.
    /// </summary>
    /// <param name="dataFormat"></param>
    /// <returns></returns>
    static ExcelCellType GetCellType(string dataFormat)
    {
        if((dataFormat.Contains("y") || dataFormat.Contains("d")) && dataFormat.Contains("h"))
            return ExcelCellType.DateTime;

        if (dataFormat.Contains("y"))
            return ExcelCellType.DateOnly;

        if (dataFormat.Contains("h") || dataFormat.Contains("m"))
            return ExcelCellType.TimeOnly;

        if (dataFormat.Contains("0") || dataFormat.Contains("#"))
            return ExcelCellType.Double;

        return ExcelCellType.Undefined;
    }
    #endregion
}

