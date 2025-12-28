using Lexerow.Core.System;
using Lexerow.Core.Tests._20_Utils;
using Lexerow.Core.Tests.Common;
using OpenExcelSdk;

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
        bool res;
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
        ExcelFile excelFile = TestExcelChecker.Open(PathExcelFilesExec + "datScriptOnExcel1.xlsx");
        Assert.IsNotNull(excelFile);

        res = TestExcelChecker.CheckCellValue(excelFile, "A2", 9);
        Assert.IsTrue(res);

        res = TestExcelChecker.CheckCellValue(excelFile, "A2", 9);
        Assert.IsTrue(res);
    }

    /// <summary>
    /// If A.Cell>B.Cell Then C.Cell=10
    /// </summary>
    [TestMethod]
    public void IfACellGreaterBCellOk()
    {
        bool res;
        LexerowCore core = new LexerowCore();
        string scriptfile = PathScriptFiles + "IfACellGreaterBCell.lxrw";

        // load the script, compile it and then execute it
        Result result = core.LoadExecScript("script", scriptfile);
        Assert.IsTrue(result.Res);

        //--check the content of excel file
        ExcelFile excelFile = TestExcelChecker.Open(PathExcelFilesExec + "IfACellGreaterBCell2.xlsx");
        Assert.IsNotNull(excelFile);

        // C2: row1, col2: 10
        res = TestExcelChecker.CheckCellValue(excelFile, "C2", 10);
        Assert.IsTrue(res);

        // C3: row2, col2: 27
        res = TestExcelChecker.CheckCellValue(excelFile, "C3", 27);
        Assert.IsTrue(res);

        // C4: row3, col2: 10
        res = TestExcelChecker.CheckCellValue(excelFile, "C4", 10);
        Assert.IsTrue(res);

        // C5: row4, col2: 10
        res = TestExcelChecker.CheckCellValue(excelFile, "C5", 10);
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
        ExcelFile excelFile = TestExcelChecker.Open(PathExcelFilesExec + "IfACellGreaterMinus7.xlsx");
        Assert.IsNotNull(excelFile);

        bool res = TestExcelChecker.CheckCellValue(excelFile, "A2", -7);
        Assert.IsTrue(res);

        res = TestExcelChecker.CheckCellValue(excelFile, "A3", -5);
        Assert.IsTrue(res);
        res = TestExcelChecker.CheckCellValue(excelFile, "A4", 4);
        Assert.IsTrue(res);
    }

    /// <summary>
    /// OnExcel ".\10-ExcelFiles\IfACellEqString.xlsx"
	///   ForEachRow
    ///     If A.Cell="Hello" Then A.Cell="Bonjour"
	///   Next
    ///   End OnExcel
    /// </summary>
    [TestMethod]
    public void IfACellEqStringOk()
    {
        LexerowCore core = new LexerowCore();
        string scriptfile = PathScriptFiles + "IfACellEqString.lxrw";

        // load the script, compile it and then execute it
        Result result = core.LoadExecScript("script", scriptfile);
        Assert.IsTrue(result.Res);

        //--check the content of excel file
        ExcelFile excelFile = TestExcelChecker.Open(PathExcelFilesExec + "IfACellEqString.xlsx");
        Assert.IsNotNull(excelFile);

        // A2: row1, col0: 10
        bool res = TestExcelChecker.CheckCellValue(excelFile, "A2", "Bonjour");
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
        ExcelFile excelFile = TestExcelChecker.Open(PathExcelFilesExec + "IfACellGreaterDouble.xlsx");
        Assert.IsNotNull(excelFile);

        // A2: row1, col0: 10
        bool res = TestExcelChecker.CheckCellValue(excelFile, "A2", 10);
        Assert.IsTrue(res);

        // A3: row2, col0: 13
        res = TestExcelChecker.CheckCellValue(excelFile, "A3", 13.1);
        Assert.IsTrue(res);

        // A4: row3, col0: 13
        res = TestExcelChecker.CheckCellValue(excelFile, "A4", 13.1);
        Assert.IsTrue(res);
    }

    /// <summary>
    /// OnExcel ".\10-ExcelFiles\onExcelManyIf.xlsx"
    ///   ForEachRow
    ///     If A.Cell <= 8 Then A.Cell=12
    ///     If B.Cell= "X" Then B.Cell= Blank
    ///     If C.Cell= 9.55 Then C.Cell= 10
    ///   Next
    /// End OnExcel
    /// </summary>
    [TestMethod]
    public void onExcelManyIfOk()
    {
        LexerowCore core = new LexerowCore();
        string scriptfile = PathScriptFiles + "onExcelManyIf.lxrw";

        // load the script, compile it and then execute it
        Result result = core.LoadExecScript("script", scriptfile);
        Assert.IsTrue(result.Res);

        //--check the content of excel file
        ExcelFile excelFile = TestExcelChecker.Open(PathExcelFilesExec + "onExcelManyIf.xlsx");
        Assert.IsNotNull(excelFile);

        //--line2: A=12, B=Y, C=10
        bool res = TestExcelChecker.CheckCellValue(excelFile, "A2", 12);
        Assert.IsTrue(res);
        res = TestExcelChecker.CheckCellValue(excelFile, "B2", "Y");
        Assert.IsTrue(res);
        res = TestExcelChecker.CheckCellValue(excelFile, "C2", 10);
        Assert.IsTrue(res);

        //--line3: A = 34, B = blank, C = 13
        res = TestExcelChecker.CheckCellValue(excelFile, "A3", 34);
        Assert.IsTrue(res);

        // B3
        res = TestExcelChecker.CheckCellValueEmpty(excelFile, "B3");
        Assert.IsTrue(res);

        // B4
        res = TestExcelChecker.CheckCellValueEmpty(excelFile, "B4");
        Assert.IsTrue(res);

        // C3
        res = TestExcelChecker.CheckCellValue(excelFile, "C3", 13);
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
        ExcelFile excelFile = TestExcelChecker.Open(PathExcelFilesExec + "onExcelManyThen.xlsx");
        Assert.IsNotNull(excelFile);

        //--line2: A=12.3, B=blank , C=Y
        bool res = TestExcelChecker.CheckCellValue(excelFile, "A2", 12.3);
        Assert.IsTrue(res);
        res = TestExcelChecker.CheckCellValueEmpty(excelFile, "B2");
        Assert.IsTrue(res);
        res = TestExcelChecker.CheckCellValue(excelFile, "C2", "Y");
        Assert.IsTrue(res);

        //--line3: A=34, B=hello, C=567
        res = TestExcelChecker.CheckCellValue(excelFile, "A3", 34);
        Assert.IsTrue(res);
        res = TestExcelChecker.CheckCellValue(excelFile, "B3", "hello");
        Assert.IsTrue(res);
        res = TestExcelChecker.CheckCellValue(excelFile, "C3", 567);
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
        ExcelFile excelFile = TestExcelChecker.Open(PathExcelFilesExec + "onExcelFirstRowValue.xlsx");
        Assert.IsNotNull(excelFile);

        //==>check some cell value

        bool res = TestExcelChecker.CheckCellValue(excelFile, "A2", 9);
        Assert.IsTrue(res);

        res = TestExcelChecker.CheckCellValue(excelFile, "A3", 12);
        Assert.IsTrue(res);

        res = TestExcelChecker.CheckCellValue(excelFile, "A4", 27);
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
        ExcelFile excelFile = TestExcelChecker.Open(PathExcelFilesExec + "onExcelFirstRowVar.xlsx");
        Assert.IsNotNull(excelFile);

        //==>check some cell value

        bool res = TestExcelChecker.CheckCellValue(excelFile, "A2", 9);
        Assert.IsTrue(res);

        res = TestExcelChecker.CheckCellValue(excelFile, "A3", 12);
        Assert.IsTrue(res);

        res = TestExcelChecker.CheckCellValue(excelFile, "A4", 27);
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
        ExcelFile excelFile = TestExcelChecker.Open(PathExcelFilesExec + "onExcelFirstRowVarVar.xlsx");
        Assert.IsNotNull(excelFile);

        //==>check some cell value

        bool res = TestExcelChecker.CheckCellValue(excelFile, "A2", 9);
        Assert.IsTrue(res);

        res = TestExcelChecker.CheckCellValue(excelFile, "A3", 12);
        Assert.IsTrue(res);

        res = TestExcelChecker.CheckCellValue(excelFile, "A4", 27);
        Assert.IsTrue(res);
    }

    /// <summary>
    /// If A.Cell <= Date(2023,11,14) Then B.Cell="yes"
    /// </summary>
    [TestMethod]
    public void OnExcelIfDateThenYesOk()
    {
        Result result;
        LexerowCore core = new LexerowCore();
        string scriptfile = PathScriptFiles + "OnExcelIfDateThenYesOk.lxrw";

        // load the script, compile it and then execute it
        result = core.LoadExecScript("script", scriptfile);
        Assert.IsTrue(result.Res);

        //--check the content of excel file
        ExcelFile excelFile = TestExcelChecker.Open(PathExcelFilesExec + "OnExcelIfDateThenYesOk.xlsx");
        Assert.IsNotNull(excelFile);

        //==>check some cell value

        bool res = TestExcelChecker.CheckCellValue(excelFile, "B2", "yes");
        Assert.IsTrue(res);

        res = TestExcelChecker.CheckCellValue(excelFile, "B3", "no");
        Assert.IsTrue(res);

        // A4: 03:40:25
        res = TestExcelChecker.CheckCellValue(excelFile, "B4", "no");
        Assert.IsTrue(res);

        // A5: special case!! 03/05/2018  12:30:45   -> compare DateTime with DateOnly -> ok
        res = TestExcelChecker.CheckCellValue(excelFile, "B5", "yes");
        Assert.IsTrue(res);
    }

    /// <summary>
    /// If A.Cell <= 10 Then B.Cell=Date(2025, 11, 30)
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
        ExcelFile excelFile = TestExcelChecker.Open(PathExcelFilesExec + "OnExcelIfThenDateOk.xlsx");
        Assert.IsNotNull(excelFile);

        //==>check some cell value
        bool res;

        // NOT modified
        res = TestExcelChecker.CheckCellValue(excelFile, "A2", 10);
        Assert.IsTrue(res);
      
        res = TestExcelChecker.CheckCellValue(excelFile, "A3", 12);
        Assert.IsTrue(res);

        // only one cell updated
        res = TestExcelChecker.CheckCellValue(excelFile, "B2", new DateOnly(2025, 11, 30));
        Assert.IsTrue(res);

        // TODO: keep it??
        res = TestExcelChecker.CheckCellStyleDataFormat(excelFile, "B2", 14);
        Assert.IsTrue(res);

        res = TestExcelChecker.CheckCellValue(excelFile, "B3", "greater");
        Assert.IsTrue(res);

        res = TestExcelChecker.CheckCellValue(excelFile, "B4", "hour");
        Assert.IsTrue(res);

        res = TestExcelChecker.CheckCellValue(excelFile, "B5", "dateTime");
        Assert.IsTrue(res);
    }

    /// <summary>
    /// If A.Cell <= Date(2023,11,14) Then A.Cell=Date(2025,3, 12)
    /// </summary>
    [TestMethod]
    public void OnExcelIfDateThenDateOk()
    {
        Result result;
        LexerowCore core = new LexerowCore();
        string scriptfile = PathScriptFiles + "OnExcelIfDateThenDateOk.lxrw";

        // load the script, compile it and then execute it
        result = core.LoadExecScript("script", scriptfile);
        Assert.IsTrue(result.Res);

        //--check the content of excel file
        ExcelFile excelFile = TestExcelChecker.Open(PathExcelFilesExec + "OnExcelIfDateThenDateOk.xlsx");
        Assert.IsNotNull(excelFile);

        //==>check some cell value

        // A2: 10/02/2019 -> 12/03/2025
        bool res = TestExcelChecker.CheckCellValue(excelFile, "A2", new DateOnly(2025, 03, 12));
        Assert.IsTrue(res);

        // A3: 03/05/2025
        res = TestExcelChecker.CheckCellValue(excelFile, "A3", new DateOnly(2025, 05, 03));
        Assert.IsTrue(res);

        // A4: 03:40:25
        res = TestExcelChecker.CheckCellValue(excelFile, "A4", new TimeOnly(03, 40, 25));
        Assert.IsTrue(res);

        // A5: 03/05/2018  12:30:45 -> 12/03/2025
        res = TestExcelChecker.CheckCellValue(excelFile, "A5", new DateOnly(2025, 03, 12));
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