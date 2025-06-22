using Lexerow.Core.System;
using Lexerow.Core.Tests._20_Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Tests.Exec;

/// <summary>
/// test exec instr ForEach Row If-Then instructions.
/// If A.Cell In [ "y", "yes", "ok" ]
/// 
/// OnExcel file
///  ForEach Row
///    If A.Cell In [ "y", "yes", "ok" ] Then A.Cell= "X"
///  
/// </summary>
[TestClass]
public class ExecOnExcelIfInItemsTests
{
    /// <summary>
    /// Test: A.Cell In [ "y", "yes", "ok" ] 
    /// 
    /// OnExcel
    ///   ForEach Row
    ///     If A.Cell In [ "y", "yes", "ok" ] Then A.Cell= "X"
    /// </summary>
    [TestMethod]
    public void OnExcelIfInStringItemsOk()
    {
        LexerowCore core = new LexerowCore();

        string fileName = @"10-Files\OnExcelIfInStringItemsOk.xlsx";

        ExecResult execResult = core.Builder.CreateInstrOpenExcel("file", fileName);

        //--A.Cell In [ "y", "yes", "ok" ] 
        List<string> listVal = ["y", "yes", "ok"];
        execResult = core.Builder.CreateInstrCompCellInItems(0, listVal, true, out InstrCompColCellInStringItems instrCompIf);
        Assert.IsTrue(execResult.Result);

        //--A.Cell= "X"
        InstrSetCellVal instrSetValThen = core.Builder.CreateInstrSetCellVal(0, "X");

        // If A.Cell In [ "y", "yes", "ok" ] Then A.Cell= "X"
        InstrIfColThen instrIfColThen;
        execResult = core.Builder.CreateInstrIfColThen(instrCompIf, instrSetValThen, out instrIfColThen);
        Assert.IsTrue(execResult.Result);

        // ForEeach Row IfColThen
        execResult = core.Builder.CreateInstrOnExcelForEachRowIfThen("file", 0, 1, instrIfColThen);
        Assert.IsTrue(execResult.Result);

        execResult = core.Exec.Execute();
        Assert.IsTrue(execResult.Result);

        // check the number of if condition which are matched
        Assert.AreEqual(2, execResult.Insights.IfCondMatchCount);
        Assert.AreEqual(4, execResult.Insights.AnalyzedDatarowCount);

        // check the content of modified excel file
        var stream = ExcelChecker.OpenExcel(fileName);
        var wb = ExcelChecker.GetWorkbook(stream);

        // A3 : yes -> X
        bool res = ExcelChecker.CheckCellValueColRow(wb, 0, 0, 2, "X");
        Assert.IsTrue(res);

        // A4 : xxx -> xxx, not modified
        res = ExcelChecker.CheckCellValueColRow(wb, 0, 0, 3, "xxx");
        Assert.IsTrue(res);

        // A5 : y -> X
        res = ExcelChecker.CheckCellValueColRow(wb, 0, 0, 4, "X");
        Assert.IsTrue(res);

        ExcelChecker.CloseExcel(stream);
    }

    /// <summary>
    /// All cell value type in col A are integer, so the if co nditino fails each time.
    /// But should finish with success.
    /// Contains 4 rows.
    /// </summary>
    [TestMethod]
    public void OnExcelIfInStringItemsErrAreNumber()
    {
        LexerowCore core = new LexerowCore();

        string fileName = @"10-Files\OnExcelIfInStringItemsErrAreNumber.xlsx";

        ExecResult execResult = core.Builder.CreateInstrOpenExcel("file", fileName);

        //--A.Cell In [ "y", "yes", "ok" ] 
        List<string> listVal = ["y", "yes", "ok"];
        execResult = core.Builder.CreateInstrCompCellInItems(0, listVal, true, out InstrCompColCellInStringItems instrCompIf);
        Assert.IsTrue(execResult.Result);

        //--A.Cell= "X"
        InstrSetCellVal instrSetValThen = core.Builder.CreateInstrSetCellVal(0, "X");

        // If A.Cell In [ "y", "yes", "ok" ] Then A.Cell= "X"
        InstrIfColThen instrIfColThen;
        execResult = core.Builder.CreateInstrIfColThen(instrCompIf, instrSetValThen, out instrIfColThen);
        Assert.IsTrue(execResult.Result);

        // ForEeach Row IfColThen
        execResult = core.Builder.CreateInstrOnExcelForEachRowIfThen("file", 0, 1, instrIfColThen);
        Assert.IsTrue(execResult.Result);

        execResult = core.Exec.Execute();

        // should finish with success, (no modification)
        Assert.IsTrue(execResult.Result);

        // should contains several warning
        Assert.AreEqual(4, execResult.ListWarning.Count);
    }
}