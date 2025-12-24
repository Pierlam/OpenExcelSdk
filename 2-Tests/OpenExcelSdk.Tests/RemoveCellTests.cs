using OpenExcelSdk.System;
using OpenExcelSdk.Tests._50_Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcelSdk.Tests;

[TestClass]
public class RemoveCellTests : TestBase
{
    [TestMethod]
    public void SetCellValueString()
    {
        bool res;
        ExcelError error;
        ExcelProcessor proc = new ExcelProcessor();

        string filename = PathFiles + "RemoveCell.xlsx";
        res = proc.Open(filename, out ExcelFile excelFile, out error);
        Assert.IsTrue(res);

        res = proc.GetSheetAt(excelFile, 0, out ExcelSheet excelSheet, out error);
        Assert.IsTrue(res);

        ExcelCell cell;
        ExcelCellValueMulti cellValueMulti;

        //--B2: already null!
        res=proc.RemoveCell(excelSheet, "B2", out error);
        Assert.IsTrue(res);

        //--B3: 
        res = proc.RemoveCell(excelSheet, "B3", out error);
        Assert.IsTrue(res);

        //--B4: 
        res = proc.RemoveCell(excelSheet, 2,4, out error);
        Assert.IsTrue(res);

        //--B5: 
        res = proc.RemoveCell(excelSheet, 2, 5, out error);
        Assert.IsTrue(res);

        // save the changes
        res = proc.Close(excelFile, out error);


        //==>check the excel content
        res = proc.Open(filename, out excelFile, out error);
        Assert.IsTrue(res);
        res = proc.GetSheetAt(excelFile, 0, out excelSheet, out error);
        Assert.IsTrue(res);

        //--B2: null
        res = proc.GetCellAt(excelSheet, 2, 2, out cell, out error);
        Assert.IsTrue(res);
        Assert.IsNull(cell);

        //--B3:
        res = proc.GetCellAt(excelSheet, 2, 3, out cell, out error);
        Assert.IsTrue(res);
        Assert.IsNull(cell);
    }
}
