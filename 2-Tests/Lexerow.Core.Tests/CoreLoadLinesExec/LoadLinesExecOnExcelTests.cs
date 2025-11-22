using Lexerow.Core.System;
using Lexerow.Core.Tests._20_Utils;
using Lexerow.Core.Tests.Common;

namespace Lexerow.Core.Tests.CoreLoadLinesExec;

/// <summary>
/// Test the script load from lines, compile and execute from the core.
/// Focus on OnExcel instruction.
/// Need to have input excel files ready.
/// </summary>
[TestClass]
public class LoadLinesExecOnExcelTests : BaseTests
{
    [TestMethod]
    public void OnExcelBasicOk()
    {
        ExecResult execResult;
        LexerowCore core = new LexerowCore();

        // create a basic script
        List<string> lines = [
            "OnExcel " + AddDblQuote(PathExcelFilesExec + "datLinesRunOnExcel1.xlsx"),
            "  ForEach Row",
            "    If A.Cell>10 Then A.Cell=10",
            "  Next",
            "End OnExcel"
            ];

        // load the script and compile it
        execResult = core.LoadLinesScript("script", lines);
        Assert.IsTrue(execResult.Result);

        execResult = core.ExecuteScript("script");
        Assert.IsTrue(execResult.Result);

        //--check the content of excel file
        var fileStream = TestExcelChecker.OpenExcel(PathExcelFilesExec + "datLinesRunOnExcel1.xlsx");
        Assert.IsNotNull(fileStream);
        var wb = TestExcelChecker.GetWorkbook(fileStream);

        // r1, c0: 9  -> not modified
        bool res = TestExcelChecker.CheckCellValue(wb, 0, 1, 0, 9);
        Assert.IsTrue(res);

        // r2, c0: 10 -> modified!
        res = TestExcelChecker.CheckCellValue(wb, 0, 2, 0, 10);
        Assert.IsTrue(res);
    }

    [TestMethod]
    public void OnExcelBasic2Ok()
    {
        ExecResult execResult;
        LexerowCore core = new LexerowCore();

        // create a basic script
        List<string> lines = [
            "OnExcel " + AddDblQuote(PathExcelFilesExec + "datLinesRunOnExcel2.xlsx"),
            "  ForEach Row",
            "    If A.Cell>10 Then A.Cell=10",
            "  Next",
            "End OnExcel"
            ];

        // load the script, compile it and execute it
        execResult = core.LoadExecLinesScript("script", lines);
        Assert.IsTrue(execResult.Result);

        //--check the content of excel file
        var fileStream = TestExcelChecker.OpenExcel(PathExcelFilesExec + "datLinesRunOnExcel2.xlsx");
        Assert.IsNotNull(fileStream);
        var wb = TestExcelChecker.GetWorkbook(fileStream);

        // r1, c0: 9  -> not modified
        bool res = TestExcelChecker.CheckCellValue(wb, 0, 1, 0, 9);
        Assert.IsTrue(res);

        // r2, c0: 10 -> modified!
        res = TestExcelChecker.CheckCellValue(wb, 0, 2, 0, 10);
        Assert.IsTrue(res);
    }

    [TestMethod]
    public void IfACellGreateBCellOk()
    {
        ExecResult execResult;
        LexerowCore core = new LexerowCore();

        // create a basic script
        List<string> lines = [
            "OnExcel " + AddDblQuote(PathExcelFilesExec + "datLinesIfACellGreaterBCell2.xlsx"),
            "  ForEach Row",
            "    If A.Cell>B.Cell Then C.Cell=10",
            "  Next",
            "End OnExcel"
            ];

        // load the script, compile it and execute it
        execResult = core.LoadExecLinesScript("script", lines);
        Assert.IsTrue(execResult.Result);

        //--check the content of excel file
        var fileStream = TestExcelChecker.OpenExcel(PathExcelFilesExec + "datLinesIfACellGreaterBCell2.xlsx");
        Assert.IsNotNull(fileStream);
        var wb = TestExcelChecker.GetWorkbook(fileStream);

        // C2: row1, col2: 10
        bool res = TestExcelChecker.CheckCellValue(wb, 0, 1, 2, 10);
        Assert.IsTrue(res);

        // C3: row2, col2: 27
        res = TestExcelChecker.CheckCellValue(wb, 0, 2, 2, 27);
        Assert.IsTrue(res);

        // C4: row3, col2: 10
        res = TestExcelChecker.CheckCellValue(wb, 0, 3, 2, 10);
        Assert.IsTrue(res);

        // C5: row4, col2: 10
        res = TestExcelChecker.CheckCellValue(wb, 0, 4, 2, 10);
        Assert.IsTrue(res);
    }
}