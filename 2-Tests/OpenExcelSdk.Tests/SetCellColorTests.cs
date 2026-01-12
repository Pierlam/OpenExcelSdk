using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using OpenExcelSdk.System;
using OpenExcelSdk.Tests._50_Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcelSdk.Tests;

[TestClass]
public class SetCellColorTests : TestBase
{
    /// <summary>
    ///  Set a color to a cell.
    ///  it is the foreground color, pattern is set to solid, bg is set to  null.
    /// </summary>
    [TestMethod]
    public void SetColorToCell()
    {
        bool res;
        ExcelProcessor proc = new ExcelProcessor();

        string filename = PathFiles + "SetCellColor.xlsx";
        ExcelFile excelFile = proc.OpenExcelFile(filename);

        ExcelSheet excelSheet = proc.GetSheetAt(excelFile, 0);

        ExcelCell cell;
        ExcelCellValue excelCellValue;

        // to check style/CellFormat creation
        var stylesPart = excelSheet.ExcelFile.WorkbookPart.WorkbookStylesPart;
        int count = stylesPart.Stylesheet.CellFormats.Elements().Count();

        //--A2: is yellow, set yellow
        ExcelCellColor cellColor = proc.SetCellColor(excelSheet, "A2", "#FFFF00");
        Assert.IsNotNull(cellColor);
        Assert.IsNull(cellColor.BgColor);
        Assert.IsNotNull(cellColor.FgColor);
        Assert.AreEqual("#FFFF00", cellColor.FgColor.Rgb);

        //--A3: blue, set to yellow
        cellColor = proc.SetCellColor(excelSheet, "A3", "#FFFF00");
        Assert.IsNotNull(cellColor);
        Assert.IsNull(cellColor.BgColor);
        Assert.IsNotNull(cellColor.FgColor);
        Assert.AreEqual("#FFFF00", cellColor.FgColor.Rgb);

        //XXX-DEBUG: save the changes
        proc.CloseExcelFile(excelFile);

        //--A4: blue, set to red
        cellColor = proc.SetCellColor(excelSheet, "A4", "#FF0000");
        Assert.IsNotNull(cellColor);
        Assert.IsNull(cellColor.BgColor);
        Assert.IsNotNull(cellColor.FgColor);
        Assert.AreEqual("#FF0000", cellColor.FgColor.Rgb);



        //--A5: null, set to yellow (cellFormat with same fill already exists)
        proc.CreateCell(excelSheet, "A5");
        cellColor = proc.SetCellColor(excelSheet, "A5", "#FFFF00");
        Assert.IsNotNull(cellColor);
        Assert.IsNull(cellColor.BgColor);
        Assert.IsNotNull(cellColor.FgColor);
        Assert.AreEqual("#FFFF00", cellColor.FgColor.Rgb);

        //--A6: null, set to green
        cellColor = proc.SetCellColor(excelSheet, "A6", "##00FF00");
        Assert.IsNotNull(cellColor);
        Assert.IsNull(cellColor.BgColor);
        Assert.IsNotNull(cellColor.FgColor);
        Assert.AreEqual("#00FF00", cellColor.FgColor.Rgb);

        // save the changes
        proc.CloseExcelFile(excelFile);

        //==>check the excel content
        excelFile = proc.OpenExcelFile(filename);
        excelSheet = proc.GetSheetAt(excelFile, 0);

        //--only one style must be created
        int countUpdate = stylesPart.Stylesheet.CellFormats.Elements().Count();
        Assert.AreEqual(count, countUpdate);

        //--A2: yellow
        cellColor= proc.GetCellColor(excelSheet, "A2");
        Assert.IsNotNull(cellColor);
        Assert.IsNull(cellColor.BgColor);
        Assert.IsNotNull(cellColor.FgColor);
        Assert.AreEqual("#FFFF00", cellColor.FgColor.Rgb);

        //--A3: yellow
        cellColor = proc.GetCellColor(excelSheet, "A3");
        Assert.IsNotNull(cellColor);
        Assert.IsNull(cellColor.BgColor);
        Assert.IsNotNull(cellColor.FgColor);
        Assert.AreEqual("#FFFF00", cellColor.FgColor.Rgb);

        //--A4: red
        cellColor = proc.GetCellColor(excelSheet, "A4");
        Assert.IsNotNull(cellColor);
        Assert.IsNull(cellColor.BgColor);
        Assert.IsNotNull(cellColor.FgColor);
        Assert.AreEqual("#FF0000", cellColor.FgColor.Rgb);

        //--A5: yellow
        cellColor = proc.GetCellColor(excelSheet, "A5");
        Assert.IsNotNull(cellColor);
        Assert.IsNull(cellColor.BgColor);
        Assert.IsNotNull(cellColor.FgColor);
        Assert.AreEqual("#FFFF00", cellColor.FgColor.Rgb);

        //--A6: green
        cellColor = proc.GetCellColor(excelSheet, "A6");
        Assert.IsNotNull(cellColor);
        Assert.IsNull(cellColor.BgColor);
        Assert.IsNotNull(cellColor.FgColor);
        Assert.AreEqual("#00FF00", cellColor.FgColor.Rgb);

    }


}
