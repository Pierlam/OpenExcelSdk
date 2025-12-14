using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenExcelSdk.System;

namespace OpenExcelSdk;

/// <summary>
/// Main class to process Excel files.
/// </summary>
public class ExcelProcessor
{
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
            error = new ExcelError(ExcelErrorCode.UnableCreateFile, ex);
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

        if(string.IsNullOrWhiteSpace(sheetName))
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

        try
        {
            Cell? cell = excelSheet.Worksheet?.Descendants<Cell>()?.Where(c => c.CellReference == addressName).FirstOrDefault();
            if (cell == null)
                // not an error
                return true;

            excelCell = new ExcelCell(excelSheet, cell);
            return true;
        }
        catch (Exception ex) 
        {
            excelError = new ExcelError(ExcelErrorCode.UnableGetCell, ex);
            return false;
        }

    }

    /// <summary>
    /// Get the type of cell value.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="excelCell"></param>
    /// <returns></returns>
    public ExcelCellValueType GetCellValueType(ExcelSheet excelSheet, ExcelCell excelCell)
    {
        if (excelCell.Cell == null) return ExcelCellValueType.Undefined;

        //--cell datatype is defined?
        if (excelCell.Cell.DataType is not null)
        {
            if (excelCell.Cell.DataType.Value == CellValues.String) return ExcelCellValueType.String;
            if (excelCell.Cell.DataType.Value == CellValues.InlineString) return ExcelCellValueType.String;
            if (excelCell.Cell.DataType.Value == CellValues.SharedString) return ExcelCellValueType.String;

            if (excelCell.Cell.DataType.Value == CellValues.Boolean) return ExcelCellValueType.Boolean;
        }

        //--cell style is defined?
        if (IsCellValueTypeDate(excelSheet, excelCell, out ExcelCellValueType valueType))
            return valueType;

        string value = excelCell.Cell.InnerText;

        // cell value is blank?
        if (value is null)return ExcelCellValueType.Undefined;

        // is it an int?
        bool res= int.TryParse(value, out int  valInt);
        if (res) return ExcelCellValueType.Integer;

        // is it a double?  cultureInfo prb: replace . by ,
        value = value.Replace('.', ',');
        res = double.TryParse(value, out double valDouble);
        if (res) return ExcelCellValueType.Double;

        // not a string, not an number, the cell value is blank
        return ExcelCellValueType.Undefined;
    }

    /// <summary>
    /// Get the value as a string.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="cell"></param>
    /// <param name="stringValue"></param>
    /// <param name="error"></param>
    /// <returns></returns>
    public bool GetCellValue(ExcelSheet excelSheet, ExcelCell excelCell, out string stringValue, out ExcelError error)
    {
        error = null;

        if(excelCell.Cell.DataType==null)
        {
            stringValue = excelCell.Cell.CellValue.InnerText;
            return true;
        }

        if (excelCell.Cell.DataType.Value == CellValues.SharedString)
        {
            // For shared strings, look up the value in the shared strings table.
            var stringTable = excelSheet.ExcelFile.WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

            // If the shared string table is missing, something is wrong. 
            // Return the index that is in the cell. 
            // Otherwise, look up the correct text in the table.
            if (stringTable is not null)
            {
                string value= excelCell.Cell.CellValue.InnerText;
                stringValue = stringTable.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
                return true;
            }
        }

        if (excelCell.Cell.DataType != null && excelCell.Cell.DataType.Value == CellValues.InlineString)
        {
            stringValue= excelCell.Cell.InlineString?.Text?.Text ?? string.Empty;
            return true;
        }

        stringValue = string.Empty;
        return true;
    }

    public bool GetCellValue(ExcelSheet excelSheet, ExcelCell excelCell, out int intValue, out ExcelError error)
    {
        intValue = 0;
        error = null;

        string value = excelCell.Cell.InnerText;

        // TODO: not tested, is it ok?
        if (excelCell.Cell.DataType == null)
        {
            bool res = int.TryParse(value, out int valInt);
            if (res)
            {
                intValue = valInt;
                return true;
            }
        }

        if (excelCell.Cell.DataType != null && excelCell.Cell.DataType.Value == CellValues.Number)
        {
            bool res = int.TryParse(value, out int valInt);
            if (res)
            {
                intValue = valInt;
                return true;
            }
        }

        error = new ExcelError(ExcelErrorCode.TypeWrong);
        return false;
    }

    public bool GetCellValue(ExcelSheet excelSheet, ExcelCell excelCell, out double doubleValue, out ExcelError error)
    {
        doubleValue = 0;
        error = null;

        string value = excelCell.Cell.InnerText;
        if (value != null)
            value = value.Replace('.', ',');
        //value = value.ToString(global::System.Globalization.CultureInfo.InvariantCulture);

        // TODO: not tested, is it ok?
        if (excelCell.Cell.DataType == null)
        {
            bool res = double.TryParse(value, out double valDouble);
            if (res)
            {
                doubleValue = valDouble;
                return true;
            }
        }

        if (excelCell.Cell.DataType != null && excelCell.Cell.DataType.Value == CellValues.Number)
        {
            bool res = double.TryParse(value, out double valDouble);
            if (res)
            {
                doubleValue = valDouble;
                return true;
            }
        }

        error = new ExcelError(ExcelErrorCode.TypeWrong);
        return false;
    }

    #endregion


    #region Create sheet, row ,cell

    public bool CreateCell(ExcelSheet excelSheet, int colIdx, int rowIdx, out ExcelCell excelCell, out ExcelError error)
    {
        string colName= ExcelUtils.GetColumnName(colIdx);
        return CreateCell(excelSheet, colName, (uint) rowIdx, out excelCell, out error);
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
            Cell cell= row.Elements<Cell>().Where(c => c.CellReference is not null && c.CellReference.Value == cellReference).First();
            excelCell= new ExcelCell(excelSheet, cell);
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

    public bool SetCellValue(ExcelSheet excelSheet, ExcelCell excelCell, string value, out ExcelError error)
    {
        error = null;
        try
        {
            WorkbookPart workbookPart = excelSheet.ExcelFile.WorkbookPart;

            // get the table
            SharedStringTablePart shareStringPart = GetOrCreateSharedStringTablePart(excelSheet.ExcelFile.WorkbookPart);

            // Insert the text into the SharedStringTablePart.
            int index = InsertSharedStringItem(value, shareStringPart);

            // Set the value of cell A1.
            excelCell.Cell.CellValue = new CellValue(index.ToString());
            excelCell.Cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);

            return true;
        }
        catch (Exception ex) 
        {
            error = new ExcelError(ExcelErrorCode.UnableSetCellValue, ex);
            return false;
        }
    }

    public bool SetCellValue(ExcelSheet excelSheet, ExcelCell excelCell, int value, out ExcelError error)
    {
        error = null;
        // Important: store as number
        excelCell.Cell.DataType = CellValues.Number;
        // Must be string in XML
        excelCell.Cell.CellValue = new CellValue(value.ToString()); 
        return true;
    }

    public bool SetCellValue(ExcelSheet excelSheet, ExcelCell excelCell, double value, out ExcelError error)
    {
        error = null;
        // Important: store as number
        excelCell.Cell.DataType = CellValues.Number;
        excelCell.Cell.CellValue = new CellValue(value.ToString(global::System.Globalization.CultureInfo.InvariantCulture));
        return true;
    }

    #endregion

    /// <summary>
    /// Get the SharedStringTablePart. If it does not exist, create a new one.
    /// </summary>
    /// <param name="excelFile"></param>
    /// <returns></returns>
    private SharedStringTablePart GetOrCreateSharedStringTablePart(WorkbookPart workbookPart)
    {
        if (workbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)        
            return workbookPart.GetPartsOfType<SharedStringTablePart>().First();
        
        return  workbookPart.AddNewPart<SharedStringTablePart>();       
    }


    private bool IsCellValueTypeDate(ExcelSheet excelSheet, ExcelCell excelCell, out ExcelCellValueType valueType)
    {
        valueType = ExcelCellValueType.Undefined;

        //--no style, not a date
        if (excelCell.Cell.StyleIndex == null)return false;

        var stylesPart = excelSheet.ExcelFile.WorkbookPart.WorkbookStylesPart;
        var cellFormat = (CellFormat)stylesPart.Stylesheet.CellFormats.ElementAt((int)excelCell.Cell.StyleIndex.Value);

        if (cellFormat.NumberFormatId == null) return false;

        uint numFmtId = cellFormat.NumberFormatId.Value;

        // Excel built-in date formats are typically between 14 and 22
        if ((numFmtId >= 14 && numFmtId <= 22) || numFmtId == 165) // 165+ are often custom date formats
        {
            // TODO: which date ???
            valueType = ExcelCellValueType.DateOnly;
            return true;
        }

        // not a date
        return false;
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
            {
                return i;
            }

            i++;
        }

        // The text does not exist in the part. Create the SharedStringItem and return its index.
        shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new Text(text)));

        return i;
    }
}
