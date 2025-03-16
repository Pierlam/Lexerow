using Lexerow.Core.System;
using Lexerow.Core.Tests._20_Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Tests.Exec;

/// <summary>
/// Execute instruction ForEachRow if-Then
/// 
/// One ForEachRow, many If-Then instr.
/// </summary>
[TestClass]
public class ExecInstrManyForEachRowIfThenTests
{
    /// <summary>
    /// possible but not efficient!
    /// scan the datarow for each of instr!!
    /// instr1: If A.Cell<10 Then B.Cell= 10
    /// instr2: If A.Cell>50 Then B.Cell= 50
    /// </summary>
    [TestMethod]
    public void TestForEachRow2IfThen()
    {
        LexerowCore core = new LexerowCore();

        string fileName = @"10-Files\Test2ForEachRowIfThen.xlsx";

        ExecResult execResult = core.Builder.CreateInstrOpenExcel("file", fileName);

        //-- A.Cell<10 
        InstrCompColCellVal instrCompIf = core.Builder.CreateInstrCompCellVal(0, InstrCompValOperator.LesserThan, 10);

        //--B.Cell= 10
        InstrSetCellVal instrSetValThen = core.Builder.CreateInstrSetCellVal(1, 10);

        // If A.Cell < 10 Then B.Cell= 10
        InstrIfColThen instrIfColThen;
        execResult = core.Builder.CreateInstrIfColThen(instrCompIf, instrSetValThen, out instrIfColThen);
        Assert.IsTrue(execResult.Result);

        // instr1: ForEach Row If A.Cell<10 Then B.Cell= 10
        execResult = core.Builder.CreateInstrForEachRowIfThen("file", 0, 1, instrIfColThen);
        Assert.IsTrue(execResult.Result);

        //-- A.Cell>50 
        instrCompIf = core.Builder.CreateInstrCompCellVal(0, InstrCompValOperator.GreaterThan, 50);

        //--B.Cell= 50
        instrSetValThen = core.Builder.CreateInstrSetCellVal(1, 50);

        // If A.Cell > 50 Then B.Cell= 50
        execResult = core.Builder.CreateInstrIfColThen(instrCompIf, instrSetValThen, out instrIfColThen);
        Assert.IsTrue(execResult.Result);

        /// instr2: ForEach Row If A.Cell>50 Then B.Cell= 50
        execResult = core.Builder.CreateInstrForEachRowIfThen("file", 0, 1, instrIfColThen);
        Assert.IsTrue(execResult.Result);

        execResult = core.Exec.Compile();
        Assert.IsTrue(execResult.Result);

        execResult = core.Exec.Execute();
        Assert.IsTrue(execResult.Result);

        // check the content of modified excel file
        var stream = ExcelChecker.OpenExcel(fileName);
        var wb = ExcelChecker.GetWorkbook(stream);
        
        // B1, =23, notmodified
        bool res = ExcelChecker.CheckCellValueColRow(wb, 0, 1, 0, 23);
        Assert.IsTrue(res);

        // B2, =10, modified
        res = ExcelChecker.CheckCellValueColRow(wb, 0, 1, 1, 10);
        Assert.IsTrue(res);

        // B3, =50, modified
        res = ExcelChecker.CheckCellValueColRow(wb, 0, 1, 2, 50);
        Assert.IsTrue(res);

        ExcelChecker.CloseExcel(stream);
    }
}
