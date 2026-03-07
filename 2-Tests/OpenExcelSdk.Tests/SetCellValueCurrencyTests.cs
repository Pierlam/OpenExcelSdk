using OpenExcelSdk.System;
using OpenExcelSdk.Tests._50_Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcelSdk.Tests;

[TestClass]
public class SetCellValueCurrencyTests : TestBase
{
    [TestMethod]
    public void SetCellValueCurrency()
    {
        bool res;
        ExcelProcessor proc = new ExcelProcessor();

        string filename = PathFiles + "SetCellValueCurrency.xlsx";
        ExcelFile excelFile = proc.OpenExcelFile(filename);

        ExcelSheet excelSheet = proc.GetSheetAt(excelFile, 0);

        ExcelCell cell;
        ExcelCellValue excelCellValue;

        // to check style/CellFormat creation
        var stylesPart = excelSheet.ExcelFile.WorkbookPart.WorkbookStylesPart;
        int count = stylesPart.Stylesheet.CellFormats.Elements().Count();

        //--B2: 123,00 € with 2 decimals, accounting format
        res = proc.SetCellValueCurrency(excelSheet, "B2", 12.34, 2, CurrencyFormat.Currency, CurrencyName.Euro);
        Assert.IsTrue(res);

        //--B3: $456,89 

        //--B4: $678,34

        // save the changes
        proc.CloseExcelFile(excelFile);

        //==>check the excel content
        excelFile = proc.OpenExcelFile(filename);
        excelSheet = proc.GetSheetAt(excelFile, 0);

        //--B2: 123,00 € with 2 decimals, accounting format
        cell = proc.GetCellAt(excelSheet, "B2");
        Assert.IsNotNull(cell);
        excelCellValue = proc.GetCellValue(excelSheet, cell);
        Assert.AreEqual(ExcelCellType.String, excelCellValue.CellType);
        Assert.AreEqual("text", excelCellValue.StringValue);

        //--B3: $456,89 

        //--B4: $678,34

        proc.CloseExcelFile(excelFile);

    }

}
