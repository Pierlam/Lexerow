using Lexerow.Core.ExcelLayer;
using Lexerow.Core.System;
using Lexerow.Core.System.Excel;
using Lexerow.Core.Tests._20_Utils;
using Lexerow.Core.Tests.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Tests.ExcelProcessor;

[TestClass]
public class ExcelProcessorSetCellStringTests : BaseTests
{
    [TestMethod]
    public void SetCellString()
    {
        // Use the lib NPOI
        ExcelProcessorNpoi proc = new ExcelProcessorNpoi();

        string fileName = PathExcelProcessorFiles + "SetCellString.xlsx";

        bool res = proc.Open(fileName, out IExcelFile excelFile, out ResultError coreError);
        Assert.IsTrue(res);

        var sheet =proc.GetSheetAt(excelFile, 0);
        IExcelCell cell;

        //--B2: null
        cell = proc.GetCellAt(sheet,1, 1);
        if (cell == null) cell = proc.CreateCell(sheet, 1, 1);
        res = proc.SetCellValue(cell, "hello");
        Assert.IsTrue(res);

        //--B3: blank
        cell = proc.GetCellAt(sheet, 2, 1);
        res = proc.SetCellValue(cell, "aze");
        Assert.IsTrue(res);

        //--B4: string
        cell = proc.GetCellAt(sheet, 3, 1);
        res = proc.SetCellValue(cell, "coucou");
        Assert.IsTrue(res);

        //--B5: int
        cell = proc.GetCellAt(sheet, 4, 1);
        res = proc.SetCellValue(cell, "bye");
        Assert.IsTrue(res);

        //--B6: double
        cell = proc.GetCellAt(sheet, 5, 1);
        res = proc.SetCellValue(cell, "houla");
        Assert.IsTrue(res);

        //--B7: date
        cell = proc.GetCellAt(sheet, 6, 1);
        res = proc.SetCellValue(cell, "aurevoir");
        Assert.IsTrue(res);

        //--B8: string+bgcolor
        cell = proc.GetCellAt(sheet, 7, 1);
        res = proc.SetCellValue(cell, "tchao");
        Assert.IsTrue(res);

        //--B9: int+bg+fg+border
        cell = proc.GetCellAt(sheet, 8, 1);
        res = proc.SetCellValue(cell, "twenty-three");
        Assert.IsTrue(res);

        //--B10: formula
        cell = proc.GetCellAt(sheet, 9, 1);
        res = proc.SetCellValue(cell, "formula");
        Assert.IsTrue(res);

        //-save the excel file
        proc.Save(excelFile);

        //--check the content of excel file
        var fileStream = TestExcelChecker.OpenExcel(PathExcelProcessorFiles + "SetCellString.xlsx");
        Assert.IsNotNull(fileStream);
        var wb = TestExcelChecker.GetWorkbook(fileStream);

        //--B2
        res = TestExcelChecker.CheckCellValue(wb, "B2", "hello");
        Assert.IsTrue(res);

        //--B3: blank
        res = TestExcelChecker.CheckCellValue(wb, "B3", "aze");
        Assert.IsTrue(res);

        res = TestExcelChecker.CheckCellValue(wb, "B9", "twenty-three");
        Assert.IsTrue(res);

        res = TestExcelChecker.CheckCellValue(wb, "B10", "formula");
        Assert.IsTrue(res);

    }

    [TestMethod]
    public void SetCellDate()
    {
        // Use the lib NPOI
        ExcelProcessorNpoi proc = new ExcelProcessorNpoi();

        string fileName = PathExcelProcessorFiles + "SetCellDate.xlsx";

        bool res = proc.Open(fileName, out IExcelFile excelFile, out ResultError coreError);
        Assert.IsTrue(res);

        var sheet = proc.GetSheetAt(excelFile, 0);
        IExcelCell cell;

        //--B2: null  ->01/03/2011
        cell = proc.GetCellAt(sheet, 1, 1);
        if (cell == null) cell = proc.CreateCell(sheet, 1, 1);
        res = proc.SetCellValue(excelFile, sheet, cell, new DateOnly(2011, 3, 1), "DD/MM/yyyy");
        Assert.IsTrue(res);

        //--B3: blank -> 08/10/2012
        cell = proc.GetCellAt(sheet, 2, 1);
        res = proc.SetCellValue(excelFile, sheet, cell, new DateOnly(2012, 10, 8), "DD/MM/yyyy");
        Assert.IsTrue(res);

        //--B7: date -> 18/05/2017
        cell = proc.GetCellAt(sheet, 6, 1);
        res = proc.SetCellValue(excelFile, sheet, cell, new DateOnly(2017, 5, 18), "m/d/yy");
        Assert.IsTrue(res);


        //-save the excel file
        proc.Save(excelFile);

        //--check the content of excel file
        var fileStream = TestExcelChecker.OpenExcel(PathExcelProcessorFiles + "SetCellDate.xlsx");
        Assert.IsNotNull(fileStream);
        var wb = TestExcelChecker.GetWorkbook(fileStream);

        //--B2
        res = TestExcelChecker.CheckCellValue(wb, "B2", "hello");
        Assert.IsTrue(res);

    }
}
