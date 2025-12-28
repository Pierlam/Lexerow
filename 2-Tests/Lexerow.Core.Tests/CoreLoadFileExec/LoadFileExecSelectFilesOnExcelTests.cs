using FakeItEasy;
using Lexerow.Core.System;
using Lexerow.Core.Tests._20_Utils;
using Lexerow.Core.Tests.Common;
using OpenExcelSdk;

namespace Lexerow.Core.Tests.CoreLoadFileExec;

/// <summary>
/// Test the script load from txt file, compile and execute from the core.
/// Focus on SelectFiles + OnExcel instruction.
/// Need to have script and input excel files ready.
/// 
/// All cases:
///   1/ OnExcel "dat.xslx"
///   2/ files=SelectFiles(), OnExcel files
///   3/ OnExcel SelectFiles()
///   4/ files="dat.xslx", OnExcel files
///   5/ files "dat*" + ".xlsx", OnExcel files  stringConcat case
/// </summary>
[TestClass]
public class LoadFileExecSelectFilesOnExcelTests : BaseTests
{
    /// <summary>
    /// Case 3: 
    /// files= SelectFiles(".\10-ExcelFiles\execSelectFilesOnExcel.xlsx")
    /// OnExcel files
    ///    ForEachRow
    ///      If A.Cell>10 Then A.Cell=10
	///    Next
    /// End OnExcel
    /// </summary>
    [TestMethod]
    public void SelectFilesOnExcelBasicOk()
    {
        Result result;
        LexerowCore core = new LexerowCore();
        string scriptfile = PathScriptFiles + "execSelectFilesOnExcel.lxrw";

        // load the script, compile it and then execute it
        result = core.LoadExecScript("script", scriptfile);
        Assert.IsTrue(result.Res);

        //--check result insights
        Assert.AreEqual(1, result.Insights.FileTotalCount);
        Assert.AreEqual(1, result.Insights.SheetTotalCount);
        Assert.AreEqual(2, result.Insights.RowTotalCount);
        Assert.AreEqual(1, result.Insights.IfCondMatchTotalCount);

        //--check the content of excel file
        ExcelFile excelFile = TestExcelChecker.Open(PathExcelFilesExec + "execSelectFilesOnExcel.xlsx");
        Assert.IsNotNull(excelFile);

        // 9 -> not modified
        bool res = TestExcelChecker.CheckCellValue(excelFile, "A2", 9);
        Assert.IsTrue(res);

        // 13 -> 10: modified!
        //res = TestExcelChecker.CheckCellValue(excelFile, 0, 2, 0, 10);
        res = TestExcelChecker.CheckCellValue(excelFile, "A3", 10);
        Assert.IsTrue(res);
    }

    /// <summary>
    /// Without explicit SelectFiles instr.
    /// 
    /// files="dat.xslx"
    /// OnExcel files
    ///    ForEachRow
    ///      If A.Cell>10 Then A.Cell=10
	///    Next
    /// End OnExcel
    /// </summary>
    [TestMethod]
    public void varfilesStringOnExcelOk()
    {
        Result result;
        LexerowCore core = new LexerowCore();
        string scriptfile = PathScriptFiles + "varfilesStringOnExcelOk.lxrw";

        // load the script, compile it and then execute it
        result = core.LoadExecScript("script", scriptfile);
        Assert.IsTrue(result.Res);

        //--check result insights
        Assert.AreEqual(1, result.Insights.FileTotalCount);
        Assert.AreEqual(1, result.Insights.SheetTotalCount);
        Assert.AreEqual(2, result.Insights.RowTotalCount);
        Assert.AreEqual(1, result.Insights.IfCondMatchTotalCount);

        //--check the content of excel file
        ExcelFile excelFile = TestExcelChecker.Open(PathExcelFilesExec + "varfilesStringOnExcelOk.xlsx");
        Assert.IsNotNull(excelFile);

        // -> not modified
        bool res = TestExcelChecker.CheckCellValue(excelFile, "A2", 9);
        Assert.IsTrue(res);

        // -> modified!
        res = TestExcelChecker.CheckCellValue(excelFile, "A3", 10);
        Assert.IsTrue(res);
    }

    /// <summary>
    /// Without explicit SelectFiles instr.
    /// 
    /// a="dat.xslx"
    /// files=a
    /// OnExcel files
    ///    ForEachRow
    ///      If A.Cell>10 Then A.Cell=10
	///    Next
    /// End OnExcel
    /// </summary>
    [TestMethod]
    public void varAfilesStringOnExcelOk()
    {
        Result result;
        LexerowCore core = new LexerowCore();
        string scriptfile = PathScriptFiles + "varAfilesStringOnExcelOk.lxrw";

        // load the script, compile it and then execute it
        result = core.LoadExecScript("script", scriptfile);
        Assert.IsTrue(result.Res);

        //--check result insights
        Assert.AreEqual(1, result.Insights.FileTotalCount);
        Assert.AreEqual(1, result.Insights.SheetTotalCount);
        Assert.AreEqual(2, result.Insights.RowTotalCount);
        Assert.AreEqual(1, result.Insights.IfCondMatchTotalCount);

        //--check the content of excel file
        ExcelFile excelFile = TestExcelChecker.Open(PathExcelFilesExec + "varAfilesStringOnExcelOk.xlsx");
        Assert.IsNotNull(excelFile);

        // -> not modified
        bool res = TestExcelChecker.CheckCellValue(excelFile, "A2", 9);
        Assert.IsTrue(res);

        // -> modified!
        res = TestExcelChecker.CheckCellValue(excelFile, "A3", 10);
        Assert.IsTrue(res);
    }

    /// <summary>
    /// files "dat*" + ".xlsx",   # stringConcat case
    /// OnExcel files
    /// </summary>
    /// TODO:
}
