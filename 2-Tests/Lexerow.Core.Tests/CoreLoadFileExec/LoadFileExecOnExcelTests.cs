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
        LexerowCore core = new LexerowCore();
        string scriptfile = PathScriptFiles + "execOnExcel1.lxrw";

        // load the script, compile it and then execute it
        Result result = core.LoadExecScript("script", scriptfile);
        Assert.IsTrue(result.Res);

        //--check result insights
        Assert.AreEqual(1, result.Insights.FileTotalCount);
        Assert.AreEqual(1, result.Insights.SheetTotalCount);
        Assert.AreEqual(2, result.Insights.RowTotalCount);
        Assert.AreEqual(1, result.Insights.IfCondMatchTotalCount);

        //--check the content of excel file
        var fileStream = TestExcelChecker.OpenExcel(PathExcelFilesExec + "datScriptOnExcel1.xlsx");
        Assert.IsNotNull(fileStream);
        var wb = TestExcelChecker.GetWorkbook(fileStream);

        // r1, c0: 9  -> not modified
        bool res = TestExcelChecker.CheckCellValue(wb, 0, "A2", 9);
        Assert.IsTrue(res);

        // r2, c0: 10 -> modified!
        res = TestExcelChecker.CheckCellValue(wb, 0, "A3", 10);
        Assert.IsTrue(res);
    }

    [TestMethod]
    public void IfACellGreaterBCellOk()
    {
        LexerowCore core = new LexerowCore();
        string scriptfile = PathScriptFiles + "IfACellGreaterBCell.lxrw";

        // load the script, compile it and then execute it
        Result result = core.LoadExecScript("script", scriptfile);
        Assert.IsTrue(result.Res);

        //--check the content of excel file
        var fileStream = TestExcelChecker.OpenExcel(PathExcelFilesExec + "IfACellGreaterBCell2.xlsx");
        Assert.IsNotNull(fileStream);
        var wb = TestExcelChecker.GetWorkbook(fileStream);

        // C2: row1, col2: 10
        bool res = TestExcelChecker.CheckCellValue(wb, 0, "C2", 10);
        Assert.IsTrue(res);

        // C3: row2, col2: 27
        res = TestExcelChecker.CheckCellValue(wb, 0, "C3", 27);
        Assert.IsTrue(res);

        // C4: row3, col2: 10
        res = TestExcelChecker.CheckCellValue(wb, 0, "C4", 10);
        Assert.IsTrue(res);

        // C5: row4, col2: 10
        res = TestExcelChecker.CheckCellValue(wb, 0, "C5", 10);
        Assert.IsTrue(res);
    }

    /// <summary>
    /// If A.Cell < -7  Then A.Cell= -7
    /// </summary>
    [TestMethod]
    public void IfACellGreaterMinus7Ok()
    {
        LexerowCore core = new LexerowCore();
        string scriptfile = PathScriptFiles + "IfACellGreaterMinus7.lxrw";

        // load the script, compile it and then execute it
        Result result = core.LoadExecScript("script", scriptfile);
        Assert.IsTrue(result.Res);

        //--check the content of excel file
        var fileStream = TestExcelChecker.OpenExcel(PathExcelFilesExec + "IfACellGreaterMinus7.xlsx");
        Assert.IsNotNull(fileStream);
        var wb = TestExcelChecker.GetWorkbook(fileStream);

        bool res = TestExcelChecker.CheckCellValue(wb, 0, "A2", -7);
        Assert.IsTrue(res);

        res = TestExcelChecker.CheckCellValue(wb, 0, "A3", -5);
        Assert.IsTrue(res);
        res = TestExcelChecker.CheckCellValue(wb, 0, "A4", 4);
        Assert.IsTrue(res);
    }

    [TestMethod]
    public void IfACellEqStringOk()
    {
        LexerowCore core = new LexerowCore();
        string scriptfile = PathScriptFiles + "IfACellEqString.lxrw";

        // load the script, compile it and then execute it
        Result result = core.LoadExecScript("script", scriptfile);
        Assert.IsTrue(result.Res);

        //--check the content of excel file
        var fileStream = TestExcelChecker.OpenExcel(PathExcelFilesExec + "IfACellEqString.xlsx");
        Assert.IsNotNull(fileStream);
        var wb = TestExcelChecker.GetWorkbook(fileStream);

        // A2: row1, col0: 10
        bool res = TestExcelChecker.CheckCellValue(wb, 0, "A2", "Bonjour");
        Assert.IsTrue(res);
    }

    [TestMethod]
    public void IfACellGreaterDoubleOk()
    {
        LexerowCore core = new LexerowCore();
        string scriptfile = PathScriptFiles + "IfACellGreaterDouble.lxrw";

        // load the script, compile it and then execute it
        Result result = core.LoadExecScript("script", scriptfile);
        Assert.IsTrue(result.Res);

        //--check the content of excel file
        var fileStream = TestExcelChecker.OpenExcel(PathExcelFilesExec + "IfACellGreaterDouble.xlsx");
        Assert.IsNotNull(fileStream);
        var wb = TestExcelChecker.GetWorkbook(fileStream);

        // A2: row1, col0: 10
        bool res = TestExcelChecker.CheckCellValue(wb, 0, "A2", 10);
        Assert.IsTrue(res);

        // A3: row2, col0: 13
        res = TestExcelChecker.CheckCellValue(wb, 0, "A3", 13.1);
        Assert.IsTrue(res);

        // A4: row3, col0: 13
        res = TestExcelChecker.CheckCellValue(wb, 0, "A4", 13.1);
        Assert.IsTrue(res);
    }

    [TestMethod]
    public void onExcelManyIfOk()
    {
        LexerowCore core = new LexerowCore();
        string scriptfile = PathScriptFiles + "onExcelManyIf.lxrw";

        // load the script, compile it and then execute it
        Result result = core.LoadExecScript("script", scriptfile);
        Assert.IsTrue(result.Res);

        //--check the content of excel file
        var fileStream = TestExcelChecker.OpenExcel(PathExcelFilesExec + "onExcelManyIf.xlsx");
        Assert.IsNotNull(fileStream);
        var wb = TestExcelChecker.GetWorkbook(fileStream);

        //--line2: A=12, B=Y, C=10
        bool res = TestExcelChecker.CheckCellValue(wb, 0, "A2", 12);
        Assert.IsTrue(res);
        res = TestExcelChecker.CheckCellValue(wb, 0, "B2", "Y");
        Assert.IsTrue(res);
        res = TestExcelChecker.CheckCellValue(wb, 0, "C2", 10);
        Assert.IsTrue(res);

        //--line3: A = 34, B = blank, C = 13
        res = TestExcelChecker.CheckCellValue(wb, 0, "A3", 34);
        Assert.IsTrue(res);
        res = TestExcelChecker.CheckCellValueBlank(wb, 0, "B3");
        Assert.IsTrue(res);
        res = TestExcelChecker.CheckCellValue(wb, 0, "C3", 13);
        Assert.IsTrue(res);
    }

    [TestMethod]
    public void onExcelManyThenOk()
    {
        LexerowCore core = new LexerowCore();
        string scriptfile = PathScriptFiles + "onExcelManyThen.lxrw";

        // load the script, compile it and then execute it
        Result result = core.LoadExecScript("script", scriptfile);
        Assert.IsTrue(result.Res);

        //--check the content of excel file
        var fileStream = TestExcelChecker.OpenExcel(PathExcelFilesExec + "onExcelManyThen.xlsx");
        Assert.IsNotNull(fileStream);
        var wb = TestExcelChecker.GetWorkbook(fileStream);

        //--line2: A=12.3, B=blank , C=Y
        bool res = TestExcelChecker.CheckCellValue(wb, 0, "A2", 12.3);
        Assert.IsTrue(res);
        res = TestExcelChecker.CheckCellValueBlank(wb, 0, "B2");
        Assert.IsTrue(res);
        res = TestExcelChecker.CheckCellValue(wb, 0, "C2", "Y");
        Assert.IsTrue(res);

        //--line3: A=34, B=hello, C=567
        res = TestExcelChecker.CheckCellValue(wb, 0, "A3", 34);
        Assert.IsTrue(res);
        res = TestExcelChecker.CheckCellValue(wb, 0, "B3", "hello");
        Assert.IsTrue(res);
        res = TestExcelChecker.CheckCellValue(wb, 0, "C3", 567);
        Assert.IsTrue(res);
    }

    [TestMethod]
    public void OnExcelFirstRowValueOk()
    {
        LexerowCore core = new LexerowCore();
        string scriptfile = PathScriptFiles + "onExcelFirstRowValue.lxrw";

        // load the script, compile it and then execute it
        Result result = core.LoadExecScript("script", scriptfile);
        Assert.IsTrue(result.Res);

        //--check the content of excel file
        var fileStream = TestExcelChecker.OpenExcel(PathExcelFilesExec + "onExcelFirstRowValue.xlsx");
        Assert.IsNotNull(fileStream);
        var wb = TestExcelChecker.GetWorkbook(fileStream);

        //==>check some cell value

        bool res = TestExcelChecker.CheckCellValue(wb, 0, "A2", 9);
        Assert.IsTrue(res);

        res = TestExcelChecker.CheckCellValue(wb, 0, "A3", 12);
        Assert.IsTrue(res);

        res = TestExcelChecker.CheckCellValue(wb, 0, "A4", 27);
        Assert.IsTrue(res);
    }

    [TestMethod]
    public void OnExcelFirstRowVarOk()
    {
        Result result;
        LexerowCore core = new LexerowCore();
        string scriptfile = PathScriptFiles + "onExcelFirstRowVar.lxrw";

        // load the script, compile it and then execute it
        result = core.LoadExecScript("script", scriptfile);
        Assert.IsTrue(result.Res);

        //--check the content of excel file
        var fileStream = TestExcelChecker.OpenExcel(PathExcelFilesExec + "onExcelFirstRowVar.xlsx");
        Assert.IsNotNull(fileStream);
        var wb = TestExcelChecker.GetWorkbook(fileStream);

        //==>check some cell value

        bool res = TestExcelChecker.CheckCellValue(wb, 0, "A2", 9);
        Assert.IsTrue(res);

        res = TestExcelChecker.CheckCellValue(wb, 0, "A3", 12);
        Assert.IsTrue(res);

        res = TestExcelChecker.CheckCellValue(wb, 0, "A4", 27);
        Assert.IsTrue(res);
    }

    [TestMethod]
    public void OnExcelFirstRowVarVarOk()
    {
        Result result;
        LexerowCore core = new LexerowCore();
        string scriptfile = PathScriptFiles + "onExcelFirstRowVarVar.lxrw";

        //=> load the script, compile it and then execute it
        result = core.LoadExecScript("script", scriptfile);
        Assert.IsTrue(result.Res);

        //=> check the content of excel file
        var fileStream = TestExcelChecker.OpenExcel(PathExcelFilesExec + "onExcelFirstRowVarVar.xlsx");
        Assert.IsNotNull(fileStream);
        var wb = TestExcelChecker.GetWorkbook(fileStream);

        //==>check some cell value

        bool res = TestExcelChecker.CheckCellValue(wb, 0, "A2", 9);
        Assert.IsTrue(res);

        res = TestExcelChecker.CheckCellValue(wb, 0, "A3", 12);
        Assert.IsTrue(res);

        res = TestExcelChecker.CheckCellValue(wb, 0, "A4", 27);
        Assert.IsTrue(res);
    }

    /// <summary>
    /// If A.Cell <= Date(2023,11,14) Then A.Cell=Date(2025,3, 12)
    /// </summary>
    [TestMethod]
    public void OnExcelIfThenDateOk()
    {
        Result result;
        LexerowCore core = new LexerowCore();
        string scriptfile = PathScriptFiles + "OnExcelIfThenDateOk.lxrw";

        // load the script, compile it and then execute it
        result = core.LoadExecScript("script", scriptfile);
        Assert.IsTrue(result.Res);

        //--check the content of excel file
        var fileStream = TestExcelChecker.OpenExcel(PathExcelFilesExec + "OnExcelIfThenDate.xlsx");
        Assert.IsNotNull(fileStream);
        var wb = TestExcelChecker.GetWorkbook(fileStream);

        //==>check some cell value

        bool res = TestExcelChecker.CheckCellValue(wb, 0, "A2", new DateOnly(2023,11,14));
        Assert.IsTrue(res);

        // 03/05/2025
        res = TestExcelChecker.CheckCellValue(wb, 0, "A3", new DateOnly(2025,05,03));
        Assert.IsTrue(res);

        // A4: 03:40:25
        res = TestExcelChecker.CheckCellValue(wb, 0, "A4", new TimeOnly(03,40,25));
        Assert.IsTrue(res);

        // A5: 03/05/2018  12:30:45
        res = TestExcelChecker.CheckCellValue(wb, 0, "A5", new DateTime(2018, 05, 03,03, 40, 25));
        Assert.IsTrue(res);
    }

    /// <summary>
    /// The path of the excel file is wrong.
    /// OnExcel instr.
    /// </summary>
    [TestMethod]
    public void OnExcelBasicPathWrong()
    {
        Result result;
        LexerowCore core = new LexerowCore();
        string scriptfile = PathScriptFiles + "execOnExcel1Err.lxrw";

        // load the script, compile it and then execute it
        result = core.LoadExecScript("script", scriptfile);
        Assert.IsFalse(result.Res);
        Assert.AreEqual(ErrorCode.ExecInstrFilePathWrong, result.ListError[0].ErrorCode);
    }
}