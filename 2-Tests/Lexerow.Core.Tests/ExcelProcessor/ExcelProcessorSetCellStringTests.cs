using Lexerow.Core.ExcelLayer;
using Lexerow.Core.System;
using Lexerow.Core.System.Excel;
using Lexerow.Core.Tests._20_Utils;
using Lexerow.Core.Tests.Common;
using NPOI.SS.Formula;
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

        //--B3: blank Bg:yellow -> 08/10/2012
        cell = proc.GetCellAt(sheet, 2, 1);
        res = proc.SetCellValue(excelFile, sheet, cell, new DateOnly(2012, 10, 8), "DD/MM/yyyy");
        Assert.IsTrue(res);

        //--B4: string -> 2013-11-17
        cell = proc.GetCellAt(sheet, 3, 1);
        res = proc.SetCellValue(excelFile, sheet, cell, new DateOnly(2013, 11, 17), "yyyy-MM-DD");
        Assert.IsTrue(res);

        //--B5: B6: B7:


        //--B7: date -> 18/05/2017
        cell = proc.GetCellAt(sheet, 6, 1);
        res = proc.SetCellValue(excelFile, sheet, cell, new DateOnly(2017, 5, 18), "m/d/yy");
        Assert.IsTrue(res);

        //--B8: string+bgcolor  -> 23/09/1987
        cell = proc.GetCellAt(sheet, 7, 1);
        res = proc.SetCellValue(excelFile, sheet, cell, new DateOnly(1987, 9, 23), "m/d/yy");
        Assert.IsTrue(res);

        //--B9: int+bg+fg+border -> 28/02/2013
        cell = proc.GetCellAt(sheet, 8, 1);
        res = proc.SetCellValue(excelFile, sheet, cell, new DateOnly(2013, 2, 28), "m/d/yy");
        Assert.IsTrue(res);

        //--B10: formula -> 23/11/2019
        cell = proc.GetCellAt(sheet, 9, 1);
        res = proc.SetCellValue(excelFile, sheet, cell, new DateOnly(2019, 11, 23), "m/d/yy");
        Assert.IsTrue(res);

        //--B11: datetime -> 17/8/2018
        cell = proc.GetCellAt(sheet, 10, 1);
        res = proc.SetCellValue(excelFile, sheet, cell, new DateOnly(2018, 8, 17), "m/d/yy");
        Assert.IsTrue(res);

        //--B12: time  -> 16/7/2015
        cell = proc.GetCellAt(sheet, 11, 1);
        res = proc.SetCellValue(excelFile, sheet, cell, new DateOnly(2015, 7, 16), "m/d/yy");
        Assert.IsTrue(res);


        //-save the excel file
        proc.Save(excelFile);

        //==>check the content of excel file
        var fileStream = TestExcelChecker.OpenExcel(PathExcelProcessorFiles + "SetCellDate.xlsx");
        Assert.IsNotNull(fileStream);
        var wb = TestExcelChecker.GetWorkbook(fileStream);

        //--B2:
        res = TestExcelChecker.CheckCellValue(wb, "B2", new DateOnly(2011, 3, 1));
        Assert.IsTrue(res);
        res = TestExcelChecker.CheckCellStyleDataFormat(wb, "B2", "DD/MM/yyyy");
        Assert.IsTrue(res);

        //--B3:
        res = TestExcelChecker.CheckCellValue(wb, "B3", new DateOnly(2012, 10, 8));
        Assert.IsTrue(res);
        res = TestExcelChecker.CheckCellStyleDataFormat(wb, "B3", "DD/MM/yyyy");
        Assert.IsTrue(res);
        // 64= yellow
        res = TestExcelChecker.CheckCellBgColor(wb, "B3", 64);
        Assert.IsTrue(res);
 
        //--B4:
        res = TestExcelChecker.CheckCellValue(wb, "B4", new DateOnly(2013, 11, 17));
        Assert.IsTrue(res);
        res = TestExcelChecker.CheckCellStyleDataFormat(wb, "B4", "yyyy-MM-DD");
        Assert.IsTrue(res);

        //--B8: string+bgcolor  -> 23/09/1987
        res = TestExcelChecker.CheckCellValue(wb, "B8", new DateOnly(1987, 9, 23));
        Assert.IsTrue(res);
        res = TestExcelChecker.CheckCellStyleDataFormat(wb, "B8", "m/d/yy");
        Assert.IsTrue(res);
        // 64= yellow
        res = TestExcelChecker.CheckCellBgColor(wb, "B8", 64);
        Assert.IsTrue(res);

        //--B9: int+bg+fg+border -> ???

    }
}
