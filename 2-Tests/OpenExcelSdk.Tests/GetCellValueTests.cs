using OpenExcelSdk.System;
using OpenExcelSdk.Tests._50_Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcelSdk.Tests;

[TestClass]
public class GetCellValueTests : TestBase
{
    [TestMethod]
    public void GetCellValueOk()
    {
        bool res;
        ExcelError error;
        ExcelProcessor proc = new ExcelProcessor();

        string filename = PathFiles + "GetCellValue.xlsx";
        res = proc.Open(filename, out ExcelFile excelFile, out error);
        Assert.IsTrue(res);

        res = proc.GetSheetAt(excelFile, 0, out ExcelSheet excelSheet, out error);
        Assert.IsTrue(res);

        ExcelCell cell;

        //--B2: null
        res = proc.GetCellAt(excelSheet, 2, 2, out cell, out error);
        Assert.IsTrue(res);
        // no error -> it's not an error! just a null cell: does not exists
        Assert.IsNull(error);
        Assert.IsNull(cell);

        //--B3: blank
        res = proc.GetCellAt(excelSheet, 2, 3, out cell, out error);
        Assert.IsTrue(res);
        Assert.IsNotNull(cell);
        var cellValueType = proc.GetCellValueType(excelSheet, cell);
        Assert.AreEqual(ExcelCellValueType.Undefined, cellValueType);

        //--B4: string: "hello"
        res = proc.GetCellAt(excelSheet, 2, 4, out cell, out error);
        Assert.IsTrue(res);
        cellValueType = proc.GetCellValueType(excelSheet, cell);
        Assert.AreEqual(ExcelCellValueType.String, cellValueType);

        res= proc.GetCellValue(excelSheet, cell, out string stringValue, out error);
        Assert.IsTrue(res);
        Assert.IsNull(error);
        Assert.AreEqual("hello", stringValue);

        //--B5: int: 12
        res = proc.GetCellAt(excelSheet, 2, 5, out cell, out error);
        Assert.IsTrue(res);
        cellValueType = proc.GetCellValueType(excelSheet, cell);
        Assert.AreEqual(ExcelCellValueType.Integer, cellValueType);

        res = proc.GetCellValue(excelSheet, cell, out int intValue, out error);
        Assert.IsTrue(res);
        Assert.IsNull(error);
        Assert.AreEqual(12, intValue);

        //--B6: double: 34,56
        res = proc.GetCellAt(excelSheet, 2, 6, out cell, out error);
        Assert.IsTrue(res);
        cellValueType = proc.GetCellValueType(excelSheet, cell);
        Assert.AreEqual(ExcelCellValueType.Double, cellValueType);

        res = proc.GetCellValue(excelSheet, cell, out double doubleValue, out error);
        Assert.IsTrue(res);
        Assert.IsNull(error);
        Assert.AreEqual(34.56, doubleValue);

        //--B7: DateOnly: 14/12/2025
        res = proc.GetCellAt(excelSheet, 2, 7, out cell, out error);
        Assert.IsTrue(res);
        cellValueType = proc.GetCellValueType(excelSheet, cell);
        Assert.AreEqual(ExcelCellValueType.DateOnly, cellValueType);

        //res = proc.GetCellValue(excelSheet, cell, out DateOnly dateOnlyValue, out error);
        //Assert.IsTrue(res);
        //Assert.IsNull(error);
        //Assert.AreEqual(34.56, doubleValue);

        //--B9: DateTime

        //--B10: TimeOnly

        //--B11: string+bgColor+fgColor+bold
        res = proc.GetCellAt(excelSheet, 2, 11, out cell, out error);
        Assert.IsTrue(res);
        cellValueType = proc.GetCellValueType(excelSheet, cell);
        Assert.AreEqual(ExcelCellValueType.String, cellValueType);

        res = proc.GetCellValue(excelSheet, cell, out stringValue, out error);
        Assert.IsTrue(res);
        Assert.IsNull(error);
        Assert.AreEqual("azerty", stringValue);

    }
}
