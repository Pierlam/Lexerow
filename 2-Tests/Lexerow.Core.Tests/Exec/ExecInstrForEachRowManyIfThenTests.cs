using Lexerow.Core.System;
using Lexerow.Core.Tests._20_Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Tests.Exec;

/// <summary>
/// test exec instr ForEach Row with many/several If-Then instructions.
/// </summary>
[TestClass]
public class ExecInstrForEachRowManyIfThenTests
{
    /// <summary>
    /// ForEach Row
    ///     If A.Cell<10 Then B.Cell= 10    ->instr1
    ///     If A.Cell>50 Then B.Cell= 50    ->instr2
    /// </summary>
    [TestMethod]
    public void TestForEachRow2IfThen()
    {
        LexerowCore core = new LexerowCore();

        string fileName = @"10-Files\TestForEachRow2IfThen.xlsx";

        ExecResult execResult = core.Builder.CreateInstrOpenExcel("file", fileName);

        //-- A.Cell<10 
        InstrCompColCellVal instrCompIf = core.Builder.CreateInstrCompCellVal(0, ValCompOperator.LesserThan, 10);

        //--B.Cell= 10
        InstrSetCellVal instrSetValThen = core.Builder.CreateInstrSetCellVal(1, 10);

        // instr1: If A.Cell<10 Then B.Cell= 10
        InstrIfColThen instrIfColThen;
        execResult = core.Builder.CreateInstrIfColThen(instrCompIf, instrSetValThen, out instrIfColThen);
        Assert.IsTrue(execResult.Result);

        //-- A.Cell>50
        InstrCompColCellVal instrCompIf2 = core.Builder.CreateInstrCompCellVal(0, ValCompOperator.GreaterThan, 50);

        //--B.Cell= 50
        InstrSetCellVal instrSetValThen2 = core.Builder.CreateInstrSetCellVal(1, 50);

        // instr1: If A.Cell>50 Then B.Cell= 50
        InstrIfColThen instrIfColThen2;
        execResult = core.Builder.CreateInstrIfColThen(instrCompIf2, instrSetValThen2, out instrIfColThen2);
        Assert.IsTrue(execResult.Result);

        // ForEach Row  -> add 2 If Col Then
        List<InstrIfColThen> listInstrIfColThen = [instrIfColThen, instrIfColThen2];

        execResult = core.Builder.CreateInstrOnExcelForEachRowIfThen("file", 0, 1, listInstrIfColThen);
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
