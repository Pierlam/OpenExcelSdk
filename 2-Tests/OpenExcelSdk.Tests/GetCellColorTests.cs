using DocumentFormat.OpenXml.Drawing;
using OpenExcelSdk.System;
using OpenExcelSdk.Tests._50_Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcelSdk.Tests;

[TestClass]
public class GetCellColorTests : TestBase
{
    [TestMethod]
    public void GetCellColors()
    {
        bool res;
        ExcelProcessor proc = new ExcelProcessor();

        string filename = PathFiles + "GetCellColors.xlsx";
        ExcelFile excelFile = proc.OpenExcelFile(filename);

        ExcelSheet excelSheet = proc.GetFirstSheet(excelFile);

        ExcelCell cell;
        ExcelCellValue excelCellValue;

        //--A2: yellow
        ExcelCellColor cellColor = proc.GetCellColor(excelSheet, "A2");
        Assert.IsNotNull(cellColor);
        Assert.AreEqual(ColorName.Yellow, cellColor.FgColor.ColorName);
        Assert.IsNull(cellColor.BgColor);

        //--A3, fg and Bg colors are null
        cellColor = proc.GetCellColor(excelSheet, "A3");
        Assert.IsNotNull(cellColor);
        Assert.IsNull(cellColor.FgColor);
        Assert.IsNull(cellColor.BgColor);


        //--A4: fg

    }
}
