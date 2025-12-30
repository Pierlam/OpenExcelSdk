using DocumentFormat.OpenXml.Spreadsheet;
using OpenExcelSdk.Tests._50_Common;

namespace OpenExcelSdk.Tests;

[TestClass]
public class GetCellTypeAndValueTests : TestBase
{
    [TestMethod]
    public void GetCellTypeAndValueSpecial()
    {
        bool res;
        ExcelProcessor proc = new ExcelProcessor();

        string filename = PathFiles + "GetCellTypeAndValueSpecial.xlsx";
        ExcelFile excelFile = proc.OpenExcelFile(filename);

        ExcelSheet excelSheet = proc.GetSheetAt(excelFile, 0);

        ExcelCell cell;
        ExcelCellValueMulti cellValueMulti;

        //--B2: null
        cell = proc.GetCellAt(excelSheet, 2, 2);
        // no cell, is null not an error!
        Assert.IsNull(cell);

        //--B3: empty, bgColor:yellow
        cell = proc.GetCellAt(excelSheet, 2, 3);
        Assert.IsNotNull(cell);
        cellValueMulti = proc.GetCellValue(excelSheet, cell);
        Assert.IsNotNull(cellValueMulti);
        // empty cell, type is undefined
        Assert.AreEqual(ExcelCellType.Undefined, cellValueMulti.CellType);
    }

    [TestMethod]
    public void GetCellTypeAndValueString()
    {
        bool res;
        ExcelProcessor proc = new ExcelProcessor();

        string filename = PathFiles + "GetCellTypeAndValueString.xlsx";
        ExcelFile excelFile = proc.OpenExcelFile(filename);

        ExcelSheet excelSheet = proc.GetSheetAt(excelFile, 0);

        ExcelCell cell;
        ExcelCellValueMulti cellValueMulti;

        //--B2: string:hello
        cell = proc.GetCellAt(excelSheet, 2, 2);
        cellValueMulti = proc.GetCellValue(excelSheet, cell);
        Assert.AreEqual(ExcelCellType.String, cellValueMulti.CellType);
        Assert.AreEqual("hello", cellValueMulti.StringValue);

        //--B3: string:wind, bgColor:yellow
        cell = proc.GetCellAt(excelSheet, 2, 3);
        cellValueMulti = proc.GetCellValue(excelSheet, cell);
        Assert.AreEqual(ExcelCellType.String, cellValueMulti.CellType);
        Assert.AreEqual("wind", cellValueMulti.StringValue);

        //--B4: string:great, border:all, thick,black
        cell = proc.GetCellAt(excelSheet, 2, 4);
        cellValueMulti = proc.GetCellValue(excelSheet, cell);
        Assert.AreEqual(ExcelCellType.String, cellValueMulti.CellType);
        Assert.AreEqual("great", cellValueMulti.StringValue);
    }

    /// <summary>
    /// https://github.com/ClosedXML/ClosedXML/wiki/NumberFormatId-Lookup-Table
    /// </summary>
    [TestMethod]
    public void GetCellTypeAndValueNumber()
    {
        bool res;
        ExcelProcessor proc = new ExcelProcessor();

        string filename = PathFiles + "GetCellTypeAndValueNumber.xlsx";
        ExcelFile excelFile = proc.OpenExcelFile(filename);

        ExcelSheet excelSheet = proc.GetSheetAt(excelFile, 0);

        ExcelCell cell;
        ExcelCellValueMulti cellValueMulti;

        //--B2: int
        cell = proc.GetCellAt(excelSheet, 2, 2);
        cellValueMulti = proc.GetCellValue(excelSheet, cell);
        Assert.AreEqual(ExcelCellType.Integer, cellValueMulti.CellType);
        Assert.AreEqual(12, cellValueMulti.IntegerValue);

        //--B3: double
        cell = proc.GetCellAt(excelSheet, 2, 3);
        cellValueMulti = proc.GetCellValue(excelSheet, cell);
        Assert.AreEqual(ExcelCellType.Double, cellValueMulti.CellType);
        Assert.AreEqual(34.56, cellValueMulti.DoubleValue);

        // --B4: double, number format: 0.00, number format id: 2
        cell = proc.GetCellAt(excelSheet, 2, 4);
        cellValueMulti = proc.GetCellValue(excelSheet, cell);
        Assert.AreEqual(ExcelCellType.Double, cellValueMulti.CellType);
        Assert.AreEqual(27.13, cellValueMulti.DoubleValue);
        Assert.AreEqual("0.00", cellValueMulti.NumberFormat);

        //--B5: double, number format: 0%, number format id: 9
        cell = proc.GetCellAt(excelSheet, 2, 5);
        cellValueMulti = proc.GetCellValue(excelSheet, cell);
        Assert.AreEqual(ExcelCellType.Double, cellValueMulti.CellType);
        // 12.5%  -> 0.125
        Assert.AreEqual(0.125, cellValueMulti.DoubleValue);

        //--B6: double+BgColor+border, 36.29
        cell = proc.GetCellAt(excelSheet, 2, 6);
        cellValueMulti = proc.GetCellValue(excelSheet, cell);
        Assert.AreEqual(ExcelCellType.Double, cellValueMulti.CellType);
        Assert.AreEqual(36.29, cellValueMulti.DoubleValue);

        //3	#,##0
        //4	#,##0.00

        //9   0 %
        //10  0.00 %

        //11  0.00E+00

        //12	# ?/?
        //13	# ??/??
    }

    /// <summary>
    /// https://github.com/ClosedXML/ClosedXML/wiki/NumberFormatId-Lookup-Table
    /// </summary>
    [TestMethod]
    public void GetCellTypeAndValueDate()
    {
        bool res;
        ExcelProcessor proc = new ExcelProcessor();

        string filename = PathFiles + "GetCellTypeAndValueDate.xlsx";
        ExcelFile excelFile = proc.OpenExcelFile(filename);

        ExcelSheet excelSheet = proc.GetSheetAt(excelFile, 0);

        ExcelCell cell;
        ExcelCellValueMulti cellValueMulti;
        CellFormat cellFormat;
        string dataFormat;
        StyleMgr styleMgr = new StyleMgr();

        //--B2: date 07/12/2019
        cell = proc.GetCellAt(excelSheet, 2, 2);
        cellValueMulti = proc.GetCellValue(excelSheet, cell);
        Assert.IsFalse(cellValueMulti.IsEmpty);
        Assert.AreEqual(ExcelCellType.DateOnly, cellValueMulti.CellType);
        Assert.AreEqual(new DateOnly(2019, 12, 7), cellValueMulti.DateOnlyValue);

        // check the style and the number format
        Assert.IsNotNull(cell.Cell.StyleIndex);
        cellFormat = proc.GetCellFormat(excelSheet, cell);
        Assert.IsNotNull(cellFormat.ApplyNumberFormat);
        // 14: built-in format: dd/MM/yyyy
        Assert.AreEqual(14, (int)cellFormat.NumberFormatId.Value);

        //--B3: datetime: 15/09/2021 12:30:45  displayed: 15/09/2021 12:30  sec are not diplayed
        cell = proc.GetCellAt(excelSheet, 2, 3);
        cellValueMulti = proc.GetCellValue(excelSheet, cell);
        Assert.AreEqual(ExcelCellType.DateTime, cellValueMulti.CellType);
        Assert.AreEqual(new DateTime(2021, 09, 15, 12, 30, 45), cellValueMulti.DateTimeValue);

        // check the style and the number format
        Assert.IsNotNull(cell.Cell.StyleIndex);
        cellFormat = proc.GetCellFormat(excelSheet, cell);
        Assert.IsNotNull(cellFormat.ApplyNumberFormat);
        // 22: built-in format: dd/MM/yyyy HH:mm
        Assert.AreEqual(22, (int)cellFormat.NumberFormatId.Value);

        //--B4: time: 09:34:56
        cell = proc.GetCellAt(excelSheet, 2, 4);
        cellValueMulti = proc.GetCellValue(excelSheet, cell);
        Assert.AreEqual(ExcelCellType.TimeOnly, cellValueMulti.CellType);
        Assert.AreEqual(new TimeOnly(09, 34, 56), cellValueMulti.TimeOnlyValue);
        // check the style and the number format
        Assert.IsNotNull(cell.Cell.StyleIndex);
        cellFormat = proc.GetCellFormat(excelSheet, cell);
        Assert.IsNotNull(cellFormat.ApplyNumberFormat);
        // 21: built-in format: HH:mm:ss
        Assert.AreEqual(21, (int)cellFormat.NumberFormatId.Value);

        //--B5: datetime, custom format: "dd/mm/yyyy\\ hh:mm:ss" -> 10/12/2025 12:34:56
        cell = proc.GetCellAt(excelSheet, 2, 5);
        cellValueMulti = proc.GetCellValue(excelSheet, cell);
        Assert.AreEqual(ExcelCellType.DateTime, cellValueMulti.CellType);
        Assert.AreEqual(new DateTime(2025, 12, 10, 12, 34, 56), cellValueMulti.DateTimeValue);
        // check the style and the number format
        Assert.IsNotNull(cell.Cell.StyleIndex);
        cellFormat = proc.GetCellFormat(excelSheet, cell);
        Assert.IsNotNull(cellFormat.ApplyNumberFormat);
        //// custom format, >164
        Assert.AreEqual(165, (int)cellFormat.NumberFormatId.Value);
        styleMgr.GetCustomNumberFormat(excelSheet, cellFormat.NumberFormatId.Value, out dataFormat);
        Assert.AreEqual("dd/mm/yyyy\\ hh:mm:ss", dataFormat);
    }

    /// <summary>
    /// https://github.com/ClosedXML/ClosedXML/wiki/NumberFormatId-Lookup-Table
    /// </summary>
    [TestMethod]
    public void GetCellTypeAndValueCustom()
    {
        bool res;
        ExcelProcessor proc = new ExcelProcessor();

        string filename = PathFiles + "GetCellTypeAndValueCustom.xlsx";
        ExcelFile excelFile = proc.OpenExcelFile(filename);

        ExcelSheet excelSheet = proc.GetSheetAt(excelFile, 0);

        ExcelCell cell;
        ExcelCellValueMulti cellValueMulti;

        //--B2: datetime, custom format: dd-MMM-yyyy HH:mm:ss -> 27/09/2025 12:34:56
        cell = proc.GetCellAt(excelSheet, 2, 2);
        cellValueMulti = proc.GetCellValue(excelSheet, cell);
        Assert.AreEqual(ExcelCellType.DateTime, cellValueMulti.CellType);
        Assert.AreEqual(new DateTime(2025, 09, 27, 12, 34, 56), cellValueMulti.DateTimeValue);
    }

    [TestMethod]
    public void GetCellTypeNullEmpty()
    {
        bool res;
        ExcelProcessor proc = new ExcelProcessor();

        string filename = PathFiles + "GetCellTypeNullEmpty.xlsx";
        ExcelFile excelFile = proc.OpenExcelFile(filename);

        ExcelSheet excelSheet = proc.GetSheetAt(excelFile, 0);

        ExcelCell cell;
        ExcelCellValueMulti cellValueMulti;

        //--B2: number, empty
        cell = proc.GetCellAt(excelSheet, "B2");
        cellValueMulti = proc.GetCellValue(excelSheet, cell);
        Assert.IsTrue(cellValueMulti.IsEmpty);
        Assert.AreEqual(ExcelCellType.Double, cellValueMulti.CellType);

        //--B3: date, bgcolor, empty
        cell = proc.GetCellAt(excelSheet, "B3");
        cellValueMulti = proc.GetCellValue(excelSheet, cell);
        Assert.IsTrue(cellValueMulti.IsEmpty);
        Assert.AreEqual(ExcelCellType.DateOnly, cellValueMulti.CellType);
    }
}