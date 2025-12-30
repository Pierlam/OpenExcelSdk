using OpenExcelSdk.Tests._50_Common;

namespace OpenExcelSdk.Tests;

[TestClass]
public class SetCellValueTests : TestBase
{
    [TestMethod]
    public void SetCellValueString()
    {
        bool res;
        ExcelError error;
        ExcelProcessor proc = new ExcelProcessor();

        string filename = PathFiles + "SetCellValueString.xlsx";
        ExcelFile excelFile = proc.OpenExcelFile(filename);

        ExcelSheet excelSheet = proc.GetSheetAt(excelFile, 0);

        ExcelCell cell;
        ExcelCellValueMulti cellValueMulti;

        // to check style/CellFormat creation
        var stylesPart = excelSheet.ExcelFile.WorkbookPart.WorkbookStylesPart;
        int count = stylesPart.Stylesheet.CellFormats.Elements().Count();

        //--B2: null
        res = proc.CreateCell(excelSheet, 2, 2, out cell, out error);
        Assert.IsTrue(res);
        // set value in a cell, if cell does not exist, it will be created
        proc.SetCellValue(excelSheet, cell, "Hello World!", out error);

        //--B3: string
        proc.SetCellValue(excelSheet, 2, 3, "montreal", out error);
        Assert.IsTrue(res);

        //--B4: string+BgColor: rain
        proc.SetCellValue(excelSheet, 2, 4, "rain", out error);
        Assert.IsTrue(res);

        //--B5: string+Border: small
        proc.SetCellValue(excelSheet, 2, 5, "small", out error);
        Assert.IsTrue(res);

        //--B6: int
        proc.SetCellValue(excelSheet, 2, 6, "other", out error);
        Assert.IsTrue(res);

        //--B7: double
        proc.SetCellValue(excelSheet, 2, 7, "green", out error);
        Assert.IsTrue(res);

        //--B8: dateOnly
        proc.SetCellValue(excelSheet, 2, 8, "mountain", out error);
        Assert.IsTrue(res);

        //--B9: double + custom format
        proc.SetCellValue(excelSheet, 2, 9, "georges", out error);
        Assert.IsTrue(res);

        //--B10: datetime + custom format
        proc.SetCellValue(excelSheet, 2, 10, "franck", out error);
        Assert.IsTrue(res);

        //--B11: formula
        proc.SetCellValue(excelSheet, 2, 11, "ferrari", out error);
        Assert.IsTrue(res);

        //--B12: formula+BgColor
        proc.SetCellValue(excelSheet, 2, 12, "fiat", out error);
        Assert.IsTrue(res);

        //--B13: date+fmt+BgColor
        proc.SetCellValue(excelSheet, 2, 13, "walker", out error);
        Assert.IsTrue(res);

        // save the changes
        proc.CloseExcelFile(excelFile);

        //==>check the excel content
        excelFile = proc.OpenExcelFile(filename);
        excelSheet = proc.GetSheetAt(excelFile, 0);

        //--only one style must be created
        int countUpdate = stylesPart.Stylesheet.CellFormats.Elements().Count();
        Assert.AreEqual(count + 1, countUpdate);

        //--B2: "Hello World!"
        res = proc.GetCellAt(excelSheet, 2, 2, out cell, out error);
        Assert.IsTrue(res);
        res = proc.GetCellTypeAndValue(excelSheet, cell, out cellValueMulti, out error);
        Assert.IsTrue(res);
        Assert.AreEqual(ExcelCellType.String, cellValueMulti.CellType);
        Assert.AreEqual("Hello World!", cellValueMulti.StringValue);

        //--B3: string
        res = proc.GetCellTypeAndValue(excelSheet, 2, 3, out cell, out cellValueMulti, out error);
        Assert.IsTrue(res);
        Assert.AreEqual(ExcelCellType.String, cellValueMulti.CellType);
        Assert.AreEqual("montreal", cellValueMulti.StringValue);

        //--B4: string+BgColor
        res = proc.GetCellTypeAndValue(excelSheet, 2, 4, out cell, out cellValueMulti, out error);
        Assert.IsTrue(res);
        Assert.AreEqual(ExcelCellType.String, cellValueMulti.CellType);
        Assert.AreEqual("rain", cellValueMulti.StringValue);

        //--B5: string+Border
        res = proc.GetCellTypeAndValue(excelSheet, 2, 5, out cell, out cellValueMulti, out error);
        Assert.IsTrue(res);
        Assert.AreEqual(ExcelCellType.String, cellValueMulti.CellType);
        Assert.AreEqual("small", cellValueMulti.StringValue);

        //--B8: string
        res = proc.GetCellTypeAndValue(excelSheet, 2, 8, out cell, out cellValueMulti, out error);
        Assert.IsTrue(res);
        Assert.AreEqual(ExcelCellType.String, cellValueMulti.CellType);
        Assert.AreEqual("mountain", cellValueMulti.StringValue);

        //--B9: string
        res = proc.GetCellTypeAndValue(excelSheet, 2, 9, out cell, out cellValueMulti, out error);
        Assert.IsTrue(res);
        Assert.AreEqual(ExcelCellType.String, cellValueMulti.CellType);
        Assert.AreEqual("georges", cellValueMulti.StringValue);

        //--B10: string
        res = proc.GetCellTypeAndValue(excelSheet, 2, 10, out cell, out cellValueMulti, out error);
        Assert.IsTrue(res);
        Assert.AreEqual(ExcelCellType.String, cellValueMulti.CellType);
        Assert.AreEqual("franck", cellValueMulti.StringValue);

        //--B11: formula
        res = proc.GetCellTypeAndValue(excelSheet, 2, 11, out cell, out cellValueMulti, out error);
        Assert.IsTrue(res);
        Assert.AreEqual(ExcelCellType.String, cellValueMulti.CellType);
        Assert.AreEqual("ferrari", cellValueMulti.StringValue);

        //--B13: was a custom datetime+fmt+BgColor, now string
        res = proc.GetCellTypeAndValue(excelSheet, 2, 13, out cell, out cellValueMulti, out error);
        Assert.IsTrue(res);
        Assert.AreEqual(ExcelCellType.String, cellValueMulti.CellType);
        Assert.AreEqual("walker", cellValueMulti.StringValue);

        // check the style and the number format
        Assert.IsNotNull(cell.Cell.StyleIndex);
        var cellFormat = proc.GetCellFormat(excelSheet, cell);

        // numberFormat must be null, no more a custom format
        Assert.IsNull(cellFormat.ApplyNumberFormat);
        Assert.IsNull(cellFormat.NumberFormatId);
    }

    [TestMethod]
    public void SetCellValueDouble()
    {
        bool res;
        ExcelError error;
        ExcelProcessor proc = new ExcelProcessor();

        string filename = PathFiles + "SetCellValueDouble.xlsx";
        ExcelFile excelFile = proc.OpenExcelFile(filename);

        ExcelSheet excelSheet = proc.GetSheetAt(excelFile, 0);

        ExcelCell cell;
        ExcelCellValueMulti cellValueMulti;

        // to check style/CellFormat creation
        var stylesPart = excelSheet.ExcelFile.WorkbookPart.WorkbookStylesPart;
        int count = stylesPart.Stylesheet.CellFormats.Elements().Count();

        //--B2: null
        proc.SetCellValue(excelSheet, 2, 2, 12.5, out error);

        //--B3: string
        proc.SetCellValue(excelSheet, 2, 3, 23.4, out error);

        //--B4:
        proc.SetCellValue(excelSheet, 2, 4, 17.2, out error);

        //--B5:
        proc.SetCellValue(excelSheet, 2, 5, 1.2, out error);

        //--B6: custom format, 21/08/1900 21:36
        proc.SetCellValue(excelSheet, 2, 6, 234.9, out error);

        //--B7:
        proc.SetCellValue(excelSheet, 2, 7, 90.1, out error);

        //--B8: custom format, 21/08/1900 21:36
        proc.SetCellValue(excelSheet, 2, 8, 456.89, out error);

        // save the changes
        proc.CloseExcelFile(excelFile);

        //==>check the excel content
        excelFile = proc.OpenExcelFile(filename);
        excelSheet = proc.GetSheetAt(excelFile, 0);

        //--only one style must be created
        int countUpdate = stylesPart.Stylesheet.CellFormats.Elements().Count();
        Assert.AreEqual(count + 1, countUpdate);

        //--B2: 12.5
        res = proc.GetCellAt(excelSheet, 2, 2, out cell, out error);
        Assert.IsTrue(res);
        res = proc.GetCellTypeAndValue(excelSheet, cell, out cellValueMulti, out error);
        Assert.IsTrue(res);
        Assert.AreEqual(ExcelCellType.Double, cellValueMulti.CellType);
        Assert.AreEqual(12.5, cellValueMulti.DoubleValue);

        //--B6: 234.9
        res = proc.GetCellAt(excelSheet, 2, 6, out cell, out error);
        Assert.IsTrue(res);
        res = proc.GetCellTypeAndValue(excelSheet, cell, out cellValueMulti, out error);
        Assert.IsTrue(res);
        Assert.AreEqual(ExcelCellType.Double, cellValueMulti.CellType);
        Assert.AreEqual(234.9, cellValueMulti.DoubleValue);
    }

    [TestMethod]
    public void SetCellValueAndFormat()
    {
        bool res;
        ExcelError error;
        ExcelProcessor proc = new ExcelProcessor();

        string filename = PathFiles + "SetCellValueAndFormat.xlsx";
        ExcelFile excelFile = proc.OpenExcelFile(filename);

        ExcelSheet excelSheet = proc.GetSheetAt(excelFile, 0);

        ExcelCell cell;
        ExcelCellValueMulti cellValueMulti;

        // to check style/CellFormat creation
        int count = proc.GetCustomNumberFormatsCount(excelSheet);

        //--B2: built-in format 2
        proc.SetCellValue(excelSheet, 2, 2, 12.5, "0.00", out error);

        //--B3: datetime custom format,  set 25.8  -> display 25.80, built-in format 2
        proc.SetCellValue(excelSheet, 2, 3, 25.8, "0.00", out error);

        //--B4: currency -> 357.20
        proc.SetCellValue(excelSheet, 2, 4, 357.2, "0.000", out error);

        //--B5: string,  -> "#,##0.00\\ \"€\""
        proc.SetCellValue(excelSheet, 2, 5, 1450, "#,##0.00\\ \"€\"", out error);

        // save the changes
        proc.CloseExcelFile(excelFile);

        //==>check the excel content
        excelFile = proc.OpenExcelFile(filename);
        excelSheet = proc.GetSheetAt(excelFile, 0);

        //--only one style must be created
        int countUpdate = proc.GetCustomNumberFormatsCount(excelSheet);
        Assert.AreEqual(count + 2, countUpdate);

        //--B2: 12.3
        res = proc.GetCellAt(excelSheet, 2, 2, out cell, out error);
        Assert.IsTrue(res);
        res = proc.GetCellTypeAndValue(excelSheet, cell, out cellValueMulti, out error);
        Assert.IsTrue(res);
        Assert.AreEqual(ExcelCellType.Double, cellValueMulti.CellType);
        Assert.AreEqual(12.5, cellValueMulti.DoubleValue);

        // check the style and the number format
        Assert.IsNotNull(cell.Cell.StyleIndex);
        var cellFormat = proc.GetCellFormat(excelSheet, cell);

        // numberFormat must be defined, is a built-in format
        Assert.IsNotNull(cellFormat.ApplyNumberFormat);
        Assert.AreEqual(2, (int)cellFormat.NumberFormatId.Value);

        //--B3: datetime custom format,  set 25.8  -> display 25.80, built-in format 2
        res = proc.GetCellAt(excelSheet, 2, 3, out cell, out error);
        Assert.IsTrue(res);
        res = proc.GetCellTypeAndValue(excelSheet, cell, out cellValueMulti, out error);
        Assert.IsTrue(res);
        Assert.AreEqual(ExcelCellType.Double, cellValueMulti.CellType);
        Assert.AreEqual(25.8, cellValueMulti.DoubleValue);

        // check the style and the number format
        Assert.IsNotNull(cell.Cell.StyleIndex);
        cellFormat = proc.GetCellFormat(excelSheet, cell);
        // numberFormat must be defined, it is a built-in format
        Assert.IsNotNull(cellFormat.ApplyNumberFormat);
        Assert.AreEqual(2, (int)cellFormat.NumberFormatId.Value);

        //--B4: 357,200
        res = proc.GetCellAt(excelSheet, 2, 4, out cell, out error);
        Assert.IsTrue(res);
        res = proc.GetCellTypeAndValue(excelSheet, cell, out cellValueMulti, out error);
        Assert.IsTrue(res);
        Assert.AreEqual(ExcelCellType.Double, cellValueMulti.CellType);
        Assert.AreEqual(357.2, cellValueMulti.DoubleValue);
        Assert.AreEqual("0.000", cellValueMulti.NumberFormat);

        // numberFormat must be defined, is a custom format > 164
        Assert.IsNotNull(cell.Cell.StyleIndex);
        cellFormat = proc.GetCellFormat(excelSheet, cell);
        Assert.IsNotNull(cellFormat.ApplyNumberFormat);
        Assert.IsTrue((int)cellFormat.NumberFormatId.Value > 163);
        StyleMgr styleMgr = new StyleMgr();
        styleMgr.GetCustomNumberFormat(excelSheet, cellFormat.NumberFormatId.Value, out string numberFormat);
        Assert.AreEqual("0.000", numberFormat);

        //--B5: int 1 450,00 €
        res = proc.GetCellAt(excelSheet, 2, 5, out cell, out error);
        Assert.IsTrue(res);
        res = proc.GetCellTypeAndValue(excelSheet, cell, out cellValueMulti, out error);
        Assert.IsTrue(res);
        // currency value is always a double
        Assert.AreEqual(ExcelCellType.Double, cellValueMulti.CellType);
        Assert.AreEqual(1450, cellValueMulti.DoubleValue);
        Assert.AreEqual("#,##0.00\\ \"€\"", cellValueMulti.NumberFormat);

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
        ExcelError error;
        ExcelProcessor proc = new ExcelProcessor();

        string filename = PathFiles + "SetCellValueDate.xlsx";
        ExcelFile excelFile = proc.OpenExcelFile(filename);

        ExcelSheet excelSheet = proc.GetSheetAt(excelFile, 0);

        ExcelCell cell;
        ExcelCellValueMulti cellValueMulti;

        // to check style/CellFormat creation
        int count = proc.GetCustomNumberFormatsCount(excelSheet);

        //--B2: 10/12/2025
        res = proc.SetCellValue(excelSheet, "B2", new DateOnly(2025, 10, 12), "d/m/yyyy", out error);
        Assert.IsTrue(res);

        //--B3: 07/05/2019
        res = proc.SetCellValue(excelSheet, "B3", new DateOnly(2019, 05, 07), "d/m/yyyy", out error);
        Assert.IsTrue(res);

        //--B4: 15/11/2020 14:30
        res = proc.SetCellValue(excelSheet, "B4", new DateTime(2020, 11, 15, 14, 30, 0), "d/m/yyyy h:mm", out error);
        Assert.IsTrue(res);

        //--B5: 02/08/2017
        res = proc.SetCellValue(excelSheet, "B5", new DateOnly(2017, 08, 02), "d/m/yyyy", out error);
        Assert.IsTrue(res);

        //--B6: 12/01/1987 11:23:45
        res = proc.SetCellValue(excelSheet, "B6", new DateTime(1987, 01, 12, 11, 23, 45), "dd/mm/yyyy\\ hh:mm:ss", out error);
        Assert.IsTrue(res);

        //--B7: 10:34:56
        res = proc.SetCellValue(excelSheet, 2, 7, new TimeOnly(10, 34, 56), "hh:mm:ss", out error);
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
        res = proc.GetCellAt(excelSheet, "B2", out cell, out error);
        Assert.IsTrue(res);
        res = proc.GetCellTypeAndValue(excelSheet, cell, out cellValueMulti, out error);
        Assert.IsTrue(res);
        Assert.AreEqual(ExcelCellType.DateOnly, cellValueMulti.CellType);
        Assert.AreEqual(new DateOnly(2025, 10, 12), cellValueMulti.DateOnlyValue);
        Assert.AreEqual(14, cellValueMulti.NumberFormatId);

        // check the style and the number format
        Assert.IsNotNull(cell.Cell.StyleIndex);
        var cellFormat = proc.GetCellFormat(excelSheet, cell);
        Assert.IsNotNull(cellFormat.ApplyNumberFormat);
        Assert.AreEqual(14, (int)cellFormat.NumberFormatId.Value);

        //--B3:
        res = proc.GetCellAt(excelSheet, 2, 3, out cell, out error);
        Assert.IsTrue(res);
        res = proc.GetCellTypeAndValue(excelSheet, cell, out cellValueMulti, out error);
        Assert.IsTrue(res);
        Assert.AreEqual(ExcelCellType.DateOnly, cellValueMulti.CellType);
        Assert.AreEqual(new DateOnly(2019, 05, 07), cellValueMulti.DateOnlyValue);
        // check the style and the number format
        Assert.IsNotNull(cell.Cell.StyleIndex);
        cellFormat = proc.GetCellFormat(excelSheet, cell);
        Assert.IsNotNull(cellFormat.ApplyNumberFormat);
        Assert.AreEqual(14, (int)cellFormat.NumberFormatId.Value);
        Assert.AreEqual(14, cellValueMulti.NumberFormatId);

        //--B4: 15/11/2020 14:30
    }

    [TestMethod]
    public void SetCellValueEmpty()
    {
        bool res;
        ExcelError error;
        ExcelProcessor proc = new ExcelProcessor();

        string filename = PathFiles + "SetCellValueEmpty.xlsx";
        ExcelFile excelFile = proc.OpenExcelFile(filename);

        ExcelSheet excelSheet = proc.GetSheetAt(excelFile, 0);

        ExcelCell cell;
        ExcelCellValueMulti cellValueMulti;

        //--B2: number
        res = proc.SetCellValueEmpty(excelSheet, 2, 2, out error);
        Assert.IsTrue(res);

        //--B3: date
        res = proc.SetCellValueEmpty(excelSheet, 2, 3, out error);
        Assert.IsTrue(res);

        // save the changes
        proc.CloseExcelFile(excelFile);

        //==>check the excel content
        excelFile = proc.OpenExcelFile(filename);
        Assert.IsTrue(res);
        excelSheet = proc.GetSheetAt(excelFile, 0);
        Assert.IsTrue(res);

        //--B2: number, empty
        res = proc.GetCellAt(excelSheet, "B2", out cell, out error);
        Assert.IsTrue(res);
        res = proc.GetCellTypeAndValue(excelSheet, cell, out cellValueMulti, out error);
        Assert.IsTrue(res);
        Assert.IsTrue(cellValueMulti.IsEmpty);
        // not able to define the type
        Assert.AreEqual(ExcelCellType.Undefined, cellValueMulti.CellType);

        //--B3: date, bgcolor, empty
        res = proc.GetCellAt(excelSheet, "B3", out cell, out error);
        Assert.IsTrue(res);
        res = proc.GetCellTypeAndValue(excelSheet, cell, out cellValueMulti, out error);
        Assert.IsTrue(res);
        Assert.IsTrue(cellValueMulti.IsEmpty);
        Assert.AreEqual(ExcelCellType.DateOnly, cellValueMulti.CellType);
    }
}