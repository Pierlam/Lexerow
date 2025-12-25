using Lexerow.Core.System;
using Lexerow.Core.Tests._20_Utils;
using Lexerow.Core.Tests.Common;
using OpenExcelSdk.System;

namespace Lexerow.Core.Tests.CoreLoadLinesExec;

/// <summary>
/// Test the script load from lines, compile and execute from the core.
/// Focus on OnExcel, If A.Cell=blank / null instruction.
/// Need to have input excel files ready.
/// </summary>
[TestClass]
public class LoadLinesExecOnExcelBlankNullTests : BaseTests
{
    /// <summary>
    /// If A.Cell=blank
    /// </summary>
    [TestMethod]
    public void IfACellEqualBlankOk()
    {
        Result result;
        LexerowCore core = new LexerowCore();

        // create a basic script
        List<string> lines = [
            "OnExcel " + AddDblQuote(PathExcelFilesExec + "datLinesACellEqualBlankOk.xlsx"),
            "  ForEach Row",
            "    If A.Cell=blank Then A.Cell=123",
            "  Next",
            "End OnExcel"
            ];

        // load the script and compile it
        result = core.LoadLinesScript("script", lines);
        Assert.IsTrue(result.Res);

        result = core.ExecuteScript("script");
        Assert.IsTrue(result.Res);

        //--check the content of excel file
        ExcelFile excelFile = TestExcelChecker.Open(PathExcelFilesExec + "datLinesACellEqualBlankOk.xlsx");
        Assert.IsNotNull(excelFile);

        // A2: r1, c0: 9  -> not modified
        bool res = TestExcelChecker.CheckCellValue(excelFile, "A2", 9);
        Assert.IsTrue(res);

        // A3: r2, c0: 123 -> modified!
        res = TestExcelChecker.CheckCellValue(excelFile, "A3", 123);
        Assert.IsTrue(res);

        // A6: r5, c0: 123 -> modified!
        res = TestExcelChecker.CheckCellValue(excelFile, "A6", 123);
        Assert.IsTrue(res);
    }

    /// <summary>
    /// OnExcel...
    ///     Then A.Cell=blank
    /// </summary>
    [TestMethod]
    public void ThenACellBlankOk()
    {
        Result result;
        LexerowCore core = new LexerowCore();

        // create a basic script
        List<string> lines = [
            "OnExcel " + AddDblQuote(PathExcelFilesExec + "datLinesThenACellBlankOk.xlsx"),
            "  ForEach Row",
            "    If A.Cell=9 Then A.Cell=blank",
            "  Next",
            "End OnExcel"
            ];

        // load the script and compile it
        result = core.LoadLinesScript("script", lines);
        Assert.IsTrue(result.Res);

        result = core.ExecuteScript("script");
        Assert.IsTrue(result.Res);

        //--check the content of excel file
        ExcelFile excelFile = TestExcelChecker.Open(PathExcelFilesExec + "datLinesThenACellBlankOk.xlsx");
        Assert.IsNotNull(excelFile);

        // A2: r1, c0: blank  -> modified
        bool res = TestExcelChecker.CheckCellValueEmpty(excelFile, "A2");
        Assert.IsTrue(res);

        // A4: r3, c0: blank  -> modified
        res = TestExcelChecker.CheckCellValueEmpty(excelFile, "A4");
        Assert.IsTrue(res);
    }
}