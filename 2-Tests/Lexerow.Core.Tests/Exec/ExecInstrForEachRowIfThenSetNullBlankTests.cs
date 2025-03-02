using Lexerow.Core.System;
using Lexerow.Core.Tests._20_Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Tests.Exec;


/// <summary>
/// test SetCellNull and SetCellValueBlank
/// </summary>
[TestClass]
public class ExecInstrForEachRowIfThenSetNullBlankTests
{
    /// <summary>
    /// Test delete a cell, remove the valueand the style.
    /// If A.Cell.Val > 10 Then B.Cell= null
    /// </summary>
    [TestMethod]
    public void SetCellNull()
    {
        LexerowCore core = new LexerowCore();

        string fileName = @"10-Files\ForIfThenSetNullBasic.xlsx";

        ExecResult execResult = core.Builder.CreateInstrOpenExcel("file", fileName);

        //--Create: A.Cell.Val > 10
        InstrCompCellVal exprCompIf = core.Builder.CreateInstrCompCellVal(0, InstrCompCellValOperator.GreaterThan, 10);

        //--Create: B.Cell =null
        InstrSetCellNull instrSetVal = core.Builder.CreateInstrSetCellNull(1);

        execResult = core.Builder.CreateInstrForEachRowIfThen("file", 0, 1, exprCompIf, instrSetVal);
        Assert.IsTrue(execResult.Result);

        execResult = core.Exec.Compile();
        Assert.IsTrue(execResult.Result);

        execResult = core.Exec.Execute();
        Assert.IsTrue(execResult.Result);

        //--check the content of modified excel file
        var stream = ExcelChecker.OpenExcel(fileName);
        var wb = ExcelChecker.GetWorkbook(stream);

        //--check, B2, becomes null  
        bool res = ExcelChecker.CheckCellNull(wb, 0, 1, 1);
        Assert.IsTrue(res);

        //--check, B3, already null
        res = ExcelChecker.CheckCellNull(wb, 0, 2, 1);
        Assert.IsTrue(res);

        //--check, B4, date becomes null
        res = ExcelChecker.CheckCellNull(wb, 0, 3, 1);
        Assert.IsTrue(res);

        ExcelChecker.CloseExcel(stream);
    }

    /// <summary>
    /// SetCellBlank -> remove the value, but keep the style
    /// If A.Cell.Val > 10 Then B.Cell= blank
    /// </summary>
    [TestMethod]
    public void SetCellValueBlank()
    {
        LexerowCore core = new LexerowCore();

        string fileName = @"10-Files\ForIfThenSetBlankBasic.xlsx";

        ExecResult execResult = core.Builder.CreateInstrOpenExcel("file", fileName);

        //--Create: A.Cell.Val > 10
        InstrCompCellVal exprCompIf = core.Builder.CreateInstrCompCellVal(0, InstrCompCellValOperator.GreaterThan, 10);

        //--Create: B.Cell =blank
        InstrSetCellValueBlank instrSetVal = core.Builder.CreateInstrSetCellValueBlank(1);

        execResult = core.Builder.CreateInstrForEachRowIfThen("file", 0, 1, exprCompIf, instrSetVal);
        Assert.IsTrue(execResult.Result);

        execResult = core.Exec.Compile();
        Assert.IsTrue(execResult.Result);

        execResult = core.Exec.Execute();
        Assert.IsTrue(execResult.Result);

        //--check the content of modified excel file
        var stream = ExcelChecker.OpenExcel(fileName);
        var wb = ExcelChecker.GetWorkbook(stream);

        //--check, B2, becomes blank  
        bool res = ExcelChecker.CheckCellValueBlank(wb, 0, 1, 1);
        Assert.IsTrue(res);

        //--check, B3, already null
        res = ExcelChecker.CheckCellNull(wb, 0, 2, 1);
        Assert.IsTrue(res);

        //--check, B4, date becomes blank
        res = ExcelChecker.CheckCellValueBlank(wb, 0, 3, 1);
        Assert.IsTrue(res);

        ExcelChecker.CloseExcel(stream);

    }
}
