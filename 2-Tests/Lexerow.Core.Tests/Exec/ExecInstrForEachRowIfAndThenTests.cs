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
/// 
/// If Cond1 And Cond2 Then 
/// 
/// exp:
/// ForEach Row
///     If A.Cell>10 And A.Cell<50 Then B.Cell= 12
/// </summary>
[TestClass]
public class ExecInstrForEachRowIfAndThenTests
{
    /// <summary>
    /// test exec instr ForEach Row If-Then instructions.
    /// 
    /// If Cond1 And Cond2 Then 
    /// 
    /// exp:
    /// ForEach Row
    ///     If A.Cell>10 And A.Cell<50 Then B.Cell= 12
    /// </summary>
    [TestMethod]
    public void TestForEachRow2IfThen()
    {
        LexerowCore core = new LexerowCore();

        string fileName = @"10-Files\TestForEachRowIfAndThen.xlsx";

        ExecResult execResult = core.ProgBuilder.CreateInstrOpenExcelParamConst("file", fileName);

        //-- A.Cell>10 
        InstrCompColCellVal instrCompIf = core.ProgBuilder.CreateInstrCompCellVal(0, ValCompOperator.GreaterThan, 10);

        //-- A.Cell<50 
        InstrCompColCellVal instrCompIf2 = core.ProgBuilder.CreateInstrCompCellVal(0, ValCompOperator.LessThan, 50);

        //--B.Cell= 12
        InstrSetCellVal instrSetValThen = core.ProgBuilder.CreateInstrSetCellVal(1, 12);

        // If A.Cell>10 And A.Cell<50 Then B.Cell= 12
        InstrIfColThen instrIfColThen;
        List<InstrRetBoolBase> listInstrCompIf = [instrCompIf, instrCompIf2];
        execResult = core.ProgBuilder.CreateInstrIfColAndThen(listInstrCompIf, instrSetValThen, out instrIfColThen);
        Assert.IsTrue(execResult.Result);

        // ForEach Row If A.Cell>10 And A.Cell<50 Then B.Cell= 12
        execResult = core.ProgBuilder.CreateInstrOnExcelForEachRowIfThen("file", 0, 1, instrIfColThen);
        Assert.IsTrue(execResult.Result);

        execResult = core.ExecuteProgram();
        Assert.IsTrue(execResult.Result);

        // check the content of modified excel file
        var stream = ExcelChecker.OpenExcel(fileName);
        var wb = ExcelChecker.GetWorkbook(stream);

        // B1, =23, notmodified
        bool res = ExcelChecker.CheckCellValueColRow(wb, 0, 1, 0, 23);
        Assert.IsTrue(res);

        // B2, =12, modified
        res = ExcelChecker.CheckCellValueColRow(wb, 0, 1, 1, 12);
        Assert.IsTrue(res);

        // B3, =12, modified
        res = ExcelChecker.CheckCellValueColRow(wb, 0, 1, 2, 12);
        Assert.IsTrue(res);

        // B4, =37, notmodified
        res = ExcelChecker.CheckCellValueColRow(wb, 0, 1, 3, 37);
        Assert.IsTrue(res);

        ExcelChecker.CloseExcel(stream);
    }
}
