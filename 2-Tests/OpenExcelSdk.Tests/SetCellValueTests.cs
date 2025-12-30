using OpenExcelSdk.Tests._50_Common;

namespace OpenExcelSdk.Tests;

[TestClass]
public class SetCellValueTests : TestBase
{
    [TestMethod]
    public void SetCellValueString()
    {
        bool res;
        ExcelProcessor proc = new ExcelProcessor();

        string filename = PathFiles + "SetCellValueString.xlsx";
        ExcelFile excelFile = proc.OpenExcelFile(filename);

        ExcelSheet excelSheet = proc.GetSheetAt(excelFile, 0);

        ExcelCell cell;
        ExcelCellValue excelCellValue;

        // to check style/CellFormat creation
        var stylesPart = excelSheet.ExcelFile.WorkbookPart.WorkbookStylesPart;
        int count = stylesPart.Stylesheet.CellFormats.Elements().Count();

        //--B2: null
        cell = proc.CreateCell(excelSheet, 2, 2);

        // set value in a cell, if cell does not exist, it will be created
        res = proc.SetCellValue(excelSheet, cell, "Hello World!");
        Assert.IsTrue(res);

        //--B3: string
        res = proc.SetCellValue(excelSheet, 2, 3, "montreal");
        Assert.IsTrue(res);

        //--B4: string+BgColor: rain
        res = proc.SetCellValue(excelSheet, 2, 4, "rain");
        Assert.IsTrue(res);

        //--B5: string+Border: small
        res = proc.SetCellValue(excelSheet, 2, 5, "small");
        Assert.IsTrue(res);

        //--B6: int
        res = proc.SetCellValue(excelSheet, 2, 6, "other");
        Assert.IsTrue(res);

        //--B7: double
        res = proc.SetCellValue(excelSheet, 2, 7, "green");
        Assert.IsTrue(res);

        //--B8: dateOnly
        res = proc.SetCellValue(excelSheet, 2, 8, "mountain");
        Assert.IsTrue(res);

        //--B9: double + custom format
        res = proc.SetCellValue(excelSheet, 2, 9, "georges");
        Assert.IsTrue(res);

        //--B10: datetime + custom format
        res = proc.SetCellValue(excelSheet, 2, 10, "franck");
        Assert.IsTrue(res);

        //--B11: formula
        res = proc.SetCellValue(excelSheet, 2, 11, "ferrari");
        Assert.IsTrue(res);

        //--B12: formula+BgColor
        res = proc.SetCellValue(excelSheet, 2, 12, "fiat");
        Assert.IsTrue(res);

        //--B13: date+fmt+BgColor
        res = proc.SetCellValue(excelSheet, 2, 13, "walker");
        Assert.IsTrue(res);

        // save the changes
        proc.CloseExcelFile(excelFile);

        //==>check the excel content
        excelFile = proc.OpenExcelFile(filename);
        excelSheet = proc.GetSheetAt(excelFile, 0);

        //--only one style must be created
        int countUpdate = stylesPart.Stylesheet.CellFormats.Elements().Count();
        Assert.AreEqual(count, countUpdate);

        //--B2: "Hello World!"
        cell = proc.GetCellAt(excelSheet, 2, 2);
        excelCellValue = proc.GetCellValue(excelSheet, cell);
        Assert.AreEqual(ExcelCellType.String, excelCellValue.CellType);
        Assert.AreEqual("Hello World!", excelCellValue.StringValue);

        //--B3: string
        excelCellValue = proc.GetCellValue(excelSheet, 2, 3);
        Assert.AreEqual(ExcelCellType.String, excelCellValue.CellType);
        Assert.AreEqual("montreal", excelCellValue.StringValue);

        //--B4: string+BgColor
        excelCellValue = proc.GetCellValue(excelSheet, 2, 4);
        Assert.AreEqual(ExcelCellType.String, excelCellValue.CellType);
        Assert.AreEqual("rain", excelCellValue.StringValue);

        //--B5: string+Border
        excelCellValue = proc.GetCellValue(excelSheet, 2, 5);
        Assert.AreEqual(ExcelCellType.String, excelCellValue.CellType);
        Assert.AreEqual("small", excelCellValue.StringValue);

        //--B8: string
        excelCellValue = proc.GetCellValue(excelSheet, 2, 8);
        Assert.AreEqual(ExcelCellType.String, excelCellValue.CellType);
        Assert.AreEqual("mountain", excelCellValue.StringValue);

        //--B9: string
        excelCellValue = proc.GetCellValue(excelSheet, 2, 9);
        Assert.AreEqual(ExcelCellType.String, excelCellValue.CellType);
        Assert.AreEqual("georges", excelCellValue.StringValue);

        //--B10: string
        excelCellValue = proc.GetCellValue(excelSheet, 2, 10);
        Assert.IsTrue(res);
        Assert.AreEqual(ExcelCellType.String, excelCellValue.CellType);
        Assert.AreEqual("franck", excelCellValue.StringValue);

        //--B11: formula
        excelCellValue = proc.GetCellValue(excelSheet, 2, 11);
        Assert.IsTrue(res);
        Assert.AreEqual(ExcelCellType.String, excelCellValue.CellType);
        Assert.AreEqual("ferrari", excelCellValue.StringValue);

        //--B13: was a custom datetime+fmt+BgColor, now string
        excelCellValue = proc.GetCellValue(excelSheet, 2, 13);
        Assert.IsTrue(res);
        Assert.AreEqual(ExcelCellType.String, excelCellValue.CellType);
        Assert.AreEqual("walker", excelCellValue.StringValue);

        // check the style and the number format
        cell = proc.GetCellAt(excelSheet, 2, 13);
        Assert.IsNotNull(cell.Cell.StyleIndex);
        var cellFormat = proc.GetCellFormat(excelSheet, cell);

        // numberFormat must be null/0, no more a custom format
        Assert.IsNull(cellFormat.ApplyNumberFormat);
        Assert.AreEqual("0", cellFormat.NumberFormatId);
    }

    [TestMethod]
    public void SetCellValueDouble()
    {
        bool res;
        ExcelProcessor proc = new ExcelProcessor();

        string filename = PathFiles + "SetCellValueDouble.xlsx";
        ExcelFile excelFile = proc.OpenExcelFile(filename);

        ExcelSheet excelSheet = proc.GetSheetAt(excelFile, 0);

        ExcelCell cell;
        ExcelCellValue excelCellValue;

        // to check style/CellFormat creation
        var stylesPart = excelSheet.ExcelFile.WorkbookPart.WorkbookStylesPart;
        int count = stylesPart.Stylesheet.CellFormats.Elements().Count();

        //--B2: null
        proc.SetCellValue(excelSheet, 2, 2, 12.5);

        //--B3: string
        proc.SetCellValue(excelSheet, 2, 3, 23.4);

        //--B4:
        proc.SetCellValue(excelSheet, 2, 4, 17.2);

        //--B5:
        proc.SetCellValue(excelSheet, 2, 5, 1.2);

        //--B6: custom format, 21/08/1900 21:36
        proc.SetCellValue(excelSheet, 2, 6, 234.9);

        //--B7:
        proc.SetCellValue(excelSheet, 2, 7, 90.1);

        //--B8: custom format, 21/08/1900 21:36
        proc.SetCellValue(excelSheet, 2, 8, 456.89);

        // save the changes
        proc.CloseExcelFile(excelFile);

        //==>check the excel content
        excelFile = proc.OpenExcelFile(filename);
        excelSheet = proc.GetSheetAt(excelFile, 0);

        //--no new style created
        int countUpdate = stylesPart.Stylesheet.CellFormats.Elements().Count();
        Assert.AreEqual(count, countUpdate);

        //--B2: 12.5
        cell = proc.GetCellAt(excelSheet, 2, 2);
        excelCellValue = proc.GetCellValue(excelSheet, cell);
        Assert.AreEqual(ExcelCellType.Double, excelCellValue.CellType);
        Assert.AreEqual(12.5, excelCellValue.DoubleValue);

        //--B6: 234.9
        cell = proc.GetCellAt(excelSheet, 2, 6);
        excelCellValue = proc.GetCellValue(excelSheet, cell);
        Assert.AreEqual(ExcelCellType.Double, excelCellValue.CellType);
        Assert.AreEqual(234.9, excelCellValue.DoubleValue);
    }

    [TestMethod]
    public void SetCellValueAndFormat()
    {
        bool res;
        ExcelProcessor proc = new ExcelProcessor();

        string filename = PathFiles + "SetCellValueAndFormat.xlsx";
        ExcelFile excelFile = proc.OpenExcelFile(filename);

        ExcelSheet excelSheet = proc.GetSheetAt(excelFile, 0);

        ExcelCell cell;
        ExcelCellValue excelCellValue;

        // to check style/CellFormat creation
        int count = proc.GetCustomNumberFormatsCount(excelSheet);

        //--B2: built-in format 2
        proc.SetCellValue(excelSheet, 2, 2, 12.5, "0.00");

        //--B3: datetime custom format,  set 25.8  -> display 25.80, built-in format 2
        proc.SetCellValue(excelSheet, 2, 3, 25.8, "0.00");

        //--B4: currency -> 357.200
        proc.SetCellValue(excelSheet, 2, 4, 357.2, "0.000");

        //--B5: string,  -> "#,##0.00\\ \"€\""
        proc.SetCellValue(excelSheet, 2, 5, 1450, "#,##0.00\\ \"€\"");

        // save the changes
        proc.CloseExcelFile(excelFile);

        //==>check the excel content
        excelFile = proc.OpenExcelFile(filename);
        excelSheet = proc.GetSheetAt(excelFile, 0);

        //--no new style created
        int countUpdate = proc.GetCustomNumberFormatsCount(excelSheet);
        Assert.AreEqual(count, countUpdate);

        //--B2: 12.3
        cell = proc.GetCellAt(excelSheet, 2, 2);
        excelCellValue = proc.GetCellValue(excelSheet, cell);
        Assert.AreEqual(ExcelCellType.Double, excelCellValue.CellType);
        Assert.AreEqual(12.5, excelCellValue.DoubleValue);

        // check the style and the number format
        Assert.IsNotNull(cell.Cell.StyleIndex);
        var cellFormat = proc.GetCellFormat(excelSheet, cell);

        // numberFormat must be defined, is a built-in format
        Assert.IsNotNull(cellFormat.ApplyNumberFormat);
        Assert.AreEqual(2, (int)cellFormat.NumberFormatId.Value);

        //--B3: datetime custom format,  set 25.8  -> display 25.80, built-in format 2
        cell = proc.GetCellAt(excelSheet, 2, 3);
        excelCellValue = proc.GetCellValue(excelSheet, cell);
        Assert.AreEqual(ExcelCellType.Double, excelCellValue.CellType);
        Assert.AreEqual(25.8, excelCellValue.DoubleValue);

        // check the style and the number format
        Assert.IsNotNull(cell.Cell.StyleIndex);
        cellFormat = proc.GetCellFormat(excelSheet, cell);
        // numberFormat must be defined, it is a built-in format
        Assert.IsNotNull(cellFormat.ApplyNumberFormat);
        Assert.AreEqual(2, (int)cellFormat.NumberFormatId.Value);

        //--B4: 357,200
        cell = proc.GetCellAt(excelSheet, 2, 4);
        excelCellValue = proc.GetCellValue(excelSheet, cell);
        Assert.AreEqual(ExcelCellType.Double, excelCellValue.CellType);
        Assert.AreEqual(357.2, excelCellValue.DoubleValue);
        Assert.AreEqual("0.000", excelCellValue.NumberFormat);

        // numberFormat must be defined, is a custom format > 164
        Assert.IsNotNull(cell.Cell.StyleIndex);
        cellFormat = proc.GetCellFormat(excelSheet, cell);
        Assert.IsNotNull(cellFormat.ApplyNumberFormat);
        Assert.IsTrue((int)cellFormat.NumberFormatId.Value > 163);
        StyleMgr styleMgr = new StyleMgr();
        styleMgr.GetCustomNumberFormat(excelSheet, cellFormat.NumberFormatId.Value, out string numberFormat);
        Assert.AreEqual("0.000", numberFormat);

        //--B5: int 1 450,00 €
        cell = proc.GetCellAt(excelSheet, 2, 5);
        excelCellValue = proc.GetCellValue(excelSheet, cell);
        // currency value is always a double
        Assert.AreEqual(ExcelCellType.Double, excelCellValue.CellType);
        Assert.AreEqual(1450, excelCellValue.DoubleValue);
        Assert.AreEqual("#,##0.00\\ \"€\"", excelCellValue.NumberFormat);

        // check the style and the number format
        Assert.IsNotNull(cell.Cell.StyleIndex);
        cellFormat = proc.GetCellFormat(excelSheet, cell);
        // numberFormat must be defined, is a custom format > 164
        Assert.IsNotNull(cellFormat.ApplyNumberFormat);
        //Assert.AreEqual(2, (int)cellFormat.NumberFormatId.Value);
        styleMgr.GetCustomNumberFormat(excelSheet, cellFormat.NumberFormatId.Value, out numberFormat);
        Assert.AreEqual("#,##0.00\\ \"€\"", numberFormat);
    }

    [TestMethod]
    public void SetCellValueDate()
    {
        bool res;
        ExcelProcessor proc = new ExcelProcessor();

        string filename = PathFiles + "SetCellValueDate.xlsx";
        ExcelFile excelFile = proc.OpenExcelFile(filename);

        ExcelSheet excelSheet = proc.GetSheetAt(excelFile, 0);

        ExcelCell cell;
        ExcelCellValue excelCellValue;

        // to check style/CellFormat creation
        int count = proc.GetCustomNumberFormatsCount(excelSheet);

        //--B2: 10/12/2025
        res = proc.SetCellValue(excelSheet, "B2", new DateOnly(2025, 10, 12), "d/m/yyyy");
        Assert.IsTrue(res);

        //--B3: 07/05/2019
        res = proc.SetCellValue(excelSheet, "B3", new DateOnly(2019, 05, 07), "d/m/yyyy");
        Assert.IsTrue(res);

        //--B4: 15/11/2020 14:30
        res = proc.SetCellValue(excelSheet, "B4", new DateTime(2020, 11, 15, 14, 30, 0), "d/m/yyyy h:mm");
        Assert.IsTrue(res);

        //--B5: 02/08/2017
        res = proc.SetCellValue(excelSheet, "B5", new DateOnly(2017, 08, 02), "d/m/yyyy");
        Assert.IsTrue(res);

        //--B6: 12/01/1987 11:23:45
        res = proc.SetCellValue(excelSheet, "B6", new DateTime(1987, 01, 12, 11, 23, 45), "dd/mm/yyyy\\ hh:mm:ss");
        Assert.IsTrue(res);

        //--B7: 10:34:56
        res = proc.SetCellValue(excelSheet, "B7", new TimeOnly(10, 34, 56), "hh:mm:ss");
        Assert.IsTrue(res);

        //ici(); 08:12:45

        // save the changes
        proc.CloseExcelFile(excelFile);

        //==>check the excel content
        excelFile = proc.OpenExcelFile(filename);
        excelSheet = proc.GetSheetAt(excelFile, 0);

        //--only one style must be created
        int countUpdate = proc.GetCustomNumberFormatsCount(excelSheet);
        //Assert.AreEqual(count + 1, countUpdate);

        //--B2:
        cell = proc.GetCellAt(excelSheet, "B2");
        excelCellValue = proc.GetCellValue(excelSheet, cell);
        Assert.AreEqual(ExcelCellType.DateOnly, excelCellValue.CellType);
        Assert.AreEqual(new DateOnly(2025, 10, 12), excelCellValue.DateOnlyValue);
        Assert.AreEqual(14, excelCellValue.NumberFormatId);

        // check the style and the number format
        Assert.IsNotNull(cell.Cell.StyleIndex);
        var cellFormat = proc.GetCellFormat(excelSheet, cell);
        Assert.IsNotNull(cellFormat.ApplyNumberFormat);
        Assert.AreEqual(14, (int)cellFormat.NumberFormatId.Value);

        //--B3:
        cell = proc.GetCellAt(excelSheet, "B3");
        excelCellValue = proc.GetCellValue(excelSheet, cell);
        Assert.AreEqual(ExcelCellType.DateOnly, excelCellValue.CellType);
        Assert.AreEqual(new DateOnly(2019, 05, 07), excelCellValue.DateOnlyValue);
        // check the style and the number format
        Assert.IsNotNull(cell.Cell.StyleIndex);
        cellFormat = proc.GetCellFormat(excelSheet, cell);
        Assert.IsNotNull(cellFormat.ApplyNumberFormat);
        Assert.AreEqual(14, (int)cellFormat.NumberFormatId.Value);
        Assert.AreEqual(14, excelCellValue.NumberFormatId);

        //--B4: 15/11/2020 14:30

        //--B5:

        //--B6:

        //--B7: 10:34:56
        cell = proc.GetCellAt(excelSheet, "B7");
        excelCellValue = proc.GetCellValue(excelSheet, cell);
        Assert.AreEqual(ExcelCellType.TimeOnly, excelCellValue.CellType);
        Assert.AreEqual(new TimeOnly(10, 34, 56), excelCellValue.TimeOnlyValue);
        // check the style and the number format
        Assert.IsNotNull(cell.Cell.StyleIndex);
        cellFormat = proc.GetCellFormat(excelSheet, cell);
        // "hh:mm:ss"
        Assert.AreEqual("1", cellFormat.ApplyNumberFormat);
        Assert.AreEqual("hh:mm:ss", excelCellValue.NumberFormat);
        // custom so >163
        Assert.IsTrue((int)cellFormat.NumberFormatId.Value > 163);
    }

    [TestMethod]
    public void SetCellValueEmpty()
    {
        bool res;
        ExcelProcessor proc = new ExcelProcessor();

        string filename = PathFiles + "SetCellValueEmpty.xlsx";
        ExcelFile excelFile = proc.OpenExcelFile(filename);

        ExcelSheet excelSheet = proc.GetSheetAt(excelFile, 0);

        ExcelCell cell;
        ExcelCellValue excelCellValue;

        //--B2: number
        res = proc.SetCellValueEmpty(excelSheet, 2, 2);
        Assert.IsTrue(res);

        //--B3: date
        res = proc.SetCellValueEmpty(excelSheet, 2, 3);
        Assert.IsTrue(res);

        // save the changes
        proc.CloseExcelFile(excelFile);

        //==>check the excel content
        excelFile = proc.OpenExcelFile(filename);
        Assert.IsTrue(res);
        excelSheet = proc.GetSheetAt(excelFile, 0);
        Assert.IsTrue(res);

        //--B2: number, empty
        cell = proc.GetCellAt(excelSheet, "B2");
        excelCellValue = proc.GetCellValue(excelSheet, cell);
        Assert.IsTrue(res);
        Assert.IsTrue(excelCellValue.IsEmpty);
        // not able to define the type
        Assert.AreEqual(ExcelCellType.Undefined, excelCellValue.CellType);

        //--B3: date, bgcolor, empty
        cell = proc.GetCellAt(excelSheet, "B3");
        Assert.IsTrue(res);
        excelCellValue = proc.GetCellValue(excelSheet, cell);
        Assert.IsTrue(res);
        Assert.IsTrue(excelCellValue.IsEmpty);
        Assert.AreEqual(ExcelCellType.DateOnly, excelCellValue.CellType);
    }
}