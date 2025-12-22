using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenExcelSdk.System;
using OpenExcelSdk.Tests._50_Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Runtime.InteropServices.JavaScript.JSType;

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
        res = proc.Open(filename, out ExcelFile excelFile, out error);
        Assert.IsTrue(res);

        res = proc.GetSheetAt(excelFile, 0, out ExcelSheet excelSheet, out error);
        Assert.IsTrue(res);

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
        res = proc.Close(excelFile, out error);


        //==>check the excel content
        res = proc.Open(filename, out excelFile, out error);
        Assert.IsTrue(res);
        res = proc.GetSheetAt(excelFile, 0, out excelSheet, out error);
        Assert.IsTrue(res);

        //--only one style must be created
        int countUpdate = stylesPart.Stylesheet.CellFormats.Elements().Count();
        Assert.AreEqual(count+1, countUpdate); 


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
        res = proc.Open(filename, out ExcelFile excelFile, out error);
        Assert.IsTrue(res);

        res = proc.GetSheetAt(excelFile, 0, out ExcelSheet excelSheet, out error);
        Assert.IsTrue(res);

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
        Assert.IsTrue(res);

        //--B7:
        proc.SetCellValue(excelSheet, 2, 7, 90.1, out error);

        //--B8: custom format, 21/08/1900 21:36
        proc.SetCellValue(excelSheet, 2, 8, 456.89, out error);

        // save the changes
        res = proc.Close(excelFile, out error);


        //==>check the excel content
        res = proc.Open(filename, out excelFile, out error);
        Assert.IsTrue(res);
        res = proc.GetSheetAt(excelFile, 0, out excelSheet, out error);
        Assert.IsTrue(res);

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
}