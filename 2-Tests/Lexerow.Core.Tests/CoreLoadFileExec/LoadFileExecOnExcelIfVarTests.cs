using Lexerow.Core.System;
using Lexerow.Core.Tests._20_Utils;
using Lexerow.Core.Tests.Common;
using OpenExcelSdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Tests.CoreLoadFileExec;

/// <summary>
/// Test OnExcel If Var condition cases.
/// </summary>
[TestClass]
public class LoadFileExecOnExcelIfVarTests : BaseTests
{
    /// <summary>
    /// a=12
    /// OnExcel ".\10-ExcelFiles\onExcelIfGreaterVarOk.xlsx"
    ///   ForEachRow
    ///     If A.Cell > a Then A.Cell=12
    ///   Next
    /// End OnExcel
    /// </summary>
    [TestMethod]
    public void onExcelIfGreaterVarOk()
    {
        LexerowCore core = new LexerowCore();
        string scriptfile = PathScriptFiles + "onExcelIfGreaterVarOk.lxrw";

        // load the script, compile it and then execute it
        Result result = core.LoadExecScript("script", scriptfile);
        Assert.IsTrue(result.Res);

        //--check the content of excel file
        ExcelFile excelFile = TestExcelChecker.Open(PathExcelFilesExec + "onExcelIfGreaterVarOk.xlsx");
        Assert.IsNotNull(excelFile);

        //--A2: 27 -> 12 
        bool res = TestExcelChecker.CheckCellValue(excelFile, "A2", 12);
        Assert.IsTrue(res);

        //--A3: 8 stay
        res = TestExcelChecker.CheckCellValue(excelFile, "A3", 8);
        Assert.IsTrue(res);

        //--A4: hello
        res = TestExcelChecker.CheckCellValue(excelFile, "A4", "hello");
        Assert.IsTrue(res);
    }

    /// <summary>
    /// a=true
    /// OnExcel ".\10-ExcelFiles\onExcelIfEqVarOk.xlsx"
    ///   ForEachRow
    ///     If a Then A.Cell=12
    ///   Next
    /// End OnExcel
    /// </summary>
    [TestMethod]
    public void OnExcelIfEqVarOk()
    {
        LexerowCore core = new LexerowCore();
        string scriptfile = PathScriptFiles + "onExcelIfEqVarOk.lxrw";

        // load the script, compile it and then execute it
        Result result = core.LoadExecScript("script", scriptfile);
        Assert.IsTrue(result.Res);

        //--check the content of excel file
        ExcelFile excelFile = TestExcelChecker.Open(PathExcelFilesExec + "onExcelIfEqVarOk.xlsx");
        Assert.IsNotNull(excelFile);

        //--A2: -> 12 
        bool res = TestExcelChecker.CheckCellValue(excelFile, "A2", 12);
        Assert.IsTrue(res);

        //--A3: -> 12
        res = TestExcelChecker.CheckCellValue(excelFile, "A3", 12);
        Assert.IsTrue(res);

        //--A4: -> 12
        res = TestExcelChecker.CheckCellValue(excelFile, "A4", 12);
        Assert.IsTrue(res);
    }

    /// <summary>
    /// a=false
    /// OnExcel ".\10-ExcelFiles\onExcelIfEqVarFalseOk.xlsx"
    ///   ForEachRow
    ///     If a Then A.Cell=12
    ///   Next
    /// End OnExcel
    /// </summary>
    [TestMethod]
    public void OnExcelIfEqVarFalseOk()
    {
        LexerowCore core = new LexerowCore();
        string scriptfile = PathScriptFiles + "onExcelIfEqVarFalseOk.lxrw";

        // load the script, compile it and then execute it
        Result result = core.LoadExecScript("script", scriptfile);
        Assert.IsTrue(result.Res);

        //--check the content of excel file
        ExcelFile excelFile = TestExcelChecker.Open(PathExcelFilesExec + "onExcelIfEqVarFalseOk.xlsx");
        Assert.IsNotNull(excelFile);

        //--A2:  
        bool res = TestExcelChecker.CheckCellValue(excelFile, "A2", 27);
        Assert.IsTrue(res);

        //--A3:
        res = TestExcelChecker.CheckCellValue(excelFile, "A3", "hello");
        Assert.IsTrue(res);

        //--A4: 10/12/2025
        res = TestExcelChecker.CheckCellValue(excelFile, "A4", new DateOnly(2025, 12,10));
        Assert.IsTrue(res);
    }

    /// <summary>
    /// a=true
    /// OnExcel ".\10-ExcelFiles\onExcelIfEqVarOk.xlsx"
    ///   ForEachRow
    ///     If (a) Then A.Cell=12
    ///   Next
    /// End OnExcel
    /// </summary>
    [TestMethod]
    public void OnExcelIfBrkEqVarVrkOk()
    {
        LexerowCore core = new LexerowCore();
        string scriptfile = PathScriptFiles + "onExcelIfBrkEqVarBrkOk.lxrw";

        // load the script, compile it and then execute it
        Result result = core.LoadExecScript("script", scriptfile);
        Assert.IsTrue(result.Res);

        //--check the content of excel file
        ExcelFile excelFile = TestExcelChecker.Open(PathExcelFilesExec + "onExcelIfBrkEqVarBrkOk.xlsx");
        Assert.IsNotNull(excelFile);

        //--A2: -> 12 
        bool res = TestExcelChecker.CheckCellValue(excelFile, "A2", 12);
        Assert.IsTrue(res);

        //--A3: -> 12
        res = TestExcelChecker.CheckCellValue(excelFile, "A3", 12);
        Assert.IsTrue(res);

        //--A4: -> 12
        res = TestExcelChecker.CheckCellValue(excelFile, "A4", 12);
        Assert.IsTrue(res);
    }

}
