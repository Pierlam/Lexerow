using Lexerow.Core.System;
using Lexerow.Core.Tests._20_Utils;
using Lexerow.Core.Tests.Common;

namespace Lexerow.Core.Tests.CoreLoadFileExec;

/// <summary>
/// Test the script load from txt file, compile and execute from the core.
/// Focus on OnExcel instruction.
/// Need to have script and input excel files ready.
/// </summary>
[TestClass]
public class LoadFileExecOnExcelTests : BaseTests
{
    [TestMethod]
    public void OnExcelBasicOk()
    {
        ExecResult execResult;
        LexerowCore core = new LexerowCore();
        string scriptfile = PathScriptFiles + "execOnExcel1.lxrw";

        // load the script, compile it and then execute it
        execResult = core.LoadExecScript("script", scriptfile);
        Assert.IsTrue(execResult.Result);

        //--check result insights
        Assert.AreEqual(1, execResult.Insights.FileTotalCount);
        Assert.AreEqual(1, execResult.Insights.SheetTotalCount);
        Assert.AreEqual(2, execResult.Insights.RowTotalCount);
        Assert.AreEqual(1, execResult.Insights.IfCondMatchTotalCount);

        //--check the content of excel file
        var fileStream = ExcelTestChecker.OpenExcel(PathExcelFilesExec + "datScriptOnExcel1.xlsx");
        Assert.IsNotNull(fileStream);
        var wb = ExcelTestChecker.GetWorkbook(fileStream);

        // r1, c0: 9  -> not modified
        bool res = ExcelTestChecker.CheckCellValue(wb, 0, 1, 0, 9);
        Assert.IsTrue(res);

        // r2, c0: 10 -> modified!
        res = ExcelTestChecker.CheckCellValue(wb, 0, 2, 0, 10);
        Assert.IsTrue(res);
    }

    [TestMethod]
    public void IfACellGreaterBCellOk()
    {
        ExecResult execResult;
        LexerowCore core = new LexerowCore();
        string scriptfile = PathScriptFiles + "IfACellGreaterBCell.lxrw";

        // load the script, compile it and then execute it
        execResult = core.LoadExecScript("script", scriptfile);
        Assert.IsTrue(execResult.Result);

        //--check the content of excel file
        var fileStream = ExcelTestChecker.OpenExcel(PathExcelFilesExec + "IfACellGreaterBCell2.xlsx");
        Assert.IsNotNull(fileStream);
        var wb = ExcelTestChecker.GetWorkbook(fileStream);

        // C2: row1, col2: 10
        bool res = ExcelTestChecker.CheckCellValue(wb, 0, 1, 2, 10);
        Assert.IsTrue(res);

        // C3: row2, col2: 27
        res = ExcelTestChecker.CheckCellValue(wb, 0, 2, 2, 27);
        Assert.IsTrue(res);

        // C4: row3, col2: 10
        res = ExcelTestChecker.CheckCellValue(wb, 0, 3, 2, 10);
        Assert.IsTrue(res);

        // C5: row4, col2: 10
        res = ExcelTestChecker.CheckCellValue(wb, 0, 4, 2, 10);
        Assert.IsTrue(res);
    }

    [TestMethod]
    public void IfACellEqStringOk()
    {
        ExecResult execResult;
        LexerowCore core = new LexerowCore();
        string scriptfile = PathScriptFiles + "IfACellEqString.lxrw";

        // load the script, compile it and then execute it
        execResult = core.LoadExecScript("script", scriptfile);
        Assert.IsTrue(execResult.Result);

        //--check the content of excel file
        var fileStream = ExcelTestChecker.OpenExcel(PathExcelFilesExec + "IfACellEqString.xlsx");
        Assert.IsNotNull(fileStream);
        var wb = ExcelTestChecker.GetWorkbook(fileStream);

        // A2: row1, col0: 10
        //bool res = ExcelTestChecker.CheckCellValue(wb, 0, 1, 0, "Bonjour");
        bool res = ExcelTestChecker.CheckCellValue(wb, 0, "A2", "Bonjour");
        Assert.IsTrue(res);
    }

    [TestMethod]
    public void IfACellGreaterDoubleOk()
    {
        ExecResult execResult;
        LexerowCore core = new LexerowCore();
        string scriptfile = PathScriptFiles + "IfACellGreaterDouble.lxrw";

        // load the script, compile it and then execute it
        execResult = core.LoadExecScript("script", scriptfile);
        Assert.IsTrue(execResult.Result);

        //--check the content of excel file
        var fileStream = ExcelTestChecker.OpenExcel(PathExcelFilesExec + "IfACellGreaterDouble.xlsx");
        Assert.IsNotNull(fileStream);
        var wb = ExcelTestChecker.GetWorkbook(fileStream);

        // A2: row1, col0: 10
        bool res = ExcelTestChecker.CheckCellValue(wb, 0, "A2", 10);
        Assert.IsTrue(res);

        // A3: row2, col0: 13
        res = ExcelTestChecker.CheckCellValue(wb, 0, "A3", 13.1);
        Assert.IsTrue(res);

        // A4: row3, col0: 13
        res = ExcelTestChecker.CheckCellValue(wb, 0, "A4", 13.1);
        Assert.IsTrue(res);
    }

    [TestMethod]
    public void onExcelManyIfOk()
    {
        ExecResult execResult;
        LexerowCore core = new LexerowCore();
        string scriptfile = PathScriptFiles + "onExcelManyIf.lxrw";

        // load the script, compile it and then execute it
        execResult = core.LoadExecScript("script", scriptfile);
        Assert.IsTrue(execResult.Result);

        //--check the content of excel file
        var fileStream = ExcelTestChecker.OpenExcel(PathExcelFilesExec + "onExcelManyIf.xlsx");
        Assert.IsNotNull(fileStream);
        var wb = ExcelTestChecker.GetWorkbook(fileStream);

        //--line2: A=12, B=Y, C=10
        bool res = ExcelTestChecker.CheckCellValue(wb, 0, "A2", 12);
        Assert.IsTrue(res);
        res = ExcelTestChecker.CheckCellValue(wb, 0, "B2", "Y");
        Assert.IsTrue(res);
        res = ExcelTestChecker.CheckCellValue(wb, 0, "C2", 10);
        Assert.IsTrue(res);

        //--line3: A = 34, B = blank, C = 13
        res = ExcelTestChecker.CheckCellValue(wb, 0, "A3", 34);
        Assert.IsTrue(res);
        res = ExcelTestChecker.CheckCellValueBlank(wb, 0, "B3");
        Assert.IsTrue(res);
        res = ExcelTestChecker.CheckCellValue(wb, 0, "C3", 13);
        Assert.IsTrue(res);
    }

    [TestMethod]
    public void onExcelManyThenOk()
    {
        ExecResult execResult;
        LexerowCore core = new LexerowCore();
        string scriptfile = PathScriptFiles + "onExcelManyThen.lxrw";

        // load the script, compile it and then execute it
        execResult = core.LoadExecScript("script", scriptfile);
        Assert.IsTrue(execResult.Result);

        //--check the content of excel file
        var fileStream = ExcelTestChecker.OpenExcel(PathExcelFilesExec + "onExcelManyThen.xlsx");
        Assert.IsNotNull(fileStream);
        var wb = ExcelTestChecker.GetWorkbook(fileStream);

        //--line2: A=12.3, B=blank , C=Y
        bool res = ExcelTestChecker.CheckCellValue(wb, 0, "A2", 12.3);
        Assert.IsTrue(res);
        res = ExcelTestChecker.CheckCellValueBlank(wb, 0, "B2");
        Assert.IsTrue(res);
        res = ExcelTestChecker.CheckCellValue(wb, 0, "C2", "Y");
        Assert.IsTrue(res);

        //--line3: A=34, B=hello, C=567
        res = ExcelTestChecker.CheckCellValue(wb, 0, "A3", 34);
        Assert.IsTrue(res);
        res = ExcelTestChecker.CheckCellValue(wb, 0, "B3", "hello");
        Assert.IsTrue(res);
        res = ExcelTestChecker.CheckCellValue(wb, 0, "C3", 567);
        Assert.IsTrue(res);
    }

    /// <summary>
    /// The path of the excel file is wrong.
    /// OnExcel instr.
    /// </summary>
    [TestMethod]
    public void OnExcelBasicPathWrong()
    {
        ExecResult execResult;
        LexerowCore core = new LexerowCore();
        string scriptfile = PathScriptFiles + "execOnExcel1Err.lxrw";

        // load the script, compile it and then execute it
        execResult = core.LoadExecScript("script", scriptfile);
        Assert.IsFalse(execResult.Result);
        Assert.AreEqual(ErrorCode.ExecInstrFilePathWrong, execResult.ListError[0].ErrorCode);
    }
}