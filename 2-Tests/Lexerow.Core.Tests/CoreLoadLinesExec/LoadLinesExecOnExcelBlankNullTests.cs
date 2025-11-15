using Lexerow.Core.System;
using Lexerow.Core.Tests._20_Utils;
using Lexerow.Core.Tests.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
        ExecResult execResult;
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
        execResult = core.LoadLinesScript("script", lines);
        Assert.IsTrue(execResult.Result);

        execResult = core.ExecuteScript("script");
        Assert.IsTrue(execResult.Result);

        //--check the content of excel file
        var fileStream = ExcelTestChecker.OpenExcel(PathExcelFilesExec + "datLinesACellEqualBlankOk.xlsx");
        Assert.IsNotNull(fileStream);
        var wb = ExcelTestChecker.GetWorkbook(fileStream);

        // A2: r1, c0: 9  -> not modified
        bool res = ExcelTestChecker.CheckCellValue(wb, 0, 1, 0, 9);
        Assert.IsTrue(res);

        // A3: r2, c0: 123 -> modified!
        res = ExcelTestChecker.CheckCellValue(wb, 0, 2, 0, 123);
        Assert.IsTrue(res);

        // A6: r5, c0: 123 -> modified!
        res = ExcelTestChecker.CheckCellValue(wb, 0, 5, 0, 123);
        Assert.IsTrue(res);
    }

    /// <summary>
    /// OnExcel...
    ///     Then A.Cell=blank
    /// </summary>
    [TestMethod]
    public void ThenACellBlankOk()
    {
        ExecResult execResult;
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
        execResult = core.LoadLinesScript("script", lines);
        Assert.IsTrue(execResult.Result);

        execResult = core.ExecuteScript("script");
        Assert.IsTrue(execResult.Result);

        //--check the content of excel file
        var fileStream = ExcelTestChecker.OpenExcel(PathExcelFilesExec + "datLinesThenACellBlankOk.xlsx");
        Assert.IsNotNull(fileStream);
        var wb = ExcelTestChecker.GetWorkbook(fileStream);

        // A2: r1, c0: blank  -> modified
        bool res = ExcelTestChecker.CheckCellValueBlank(wb, 0, 1, 0);
        Assert.IsTrue(res);

        // A4: r3, c0: blank  -> modified
        res = ExcelTestChecker.CheckCellValueBlank(wb, 0, 3, 0);
        Assert.IsTrue(res);
    }

}
