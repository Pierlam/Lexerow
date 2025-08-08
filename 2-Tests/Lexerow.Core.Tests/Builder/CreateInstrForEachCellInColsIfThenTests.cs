using Lexerow.Core.System;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Tests.Builder;


[TestClass]
public  class CreateInstrForEachCellInColsIfThenTests
{
    /// <summary>
    /// Check the creation of the instr ForEachCellInColsIfThen.
    /// </summary>
    [TestMethod]
    public void CreateInstrForEachCellInColsIfThenOk()
    {
        LexerowCore core = new LexerowCore();

        string fileName = @"..\..\..\10-Files\TestCoreOpenClose.xlsx";

        ExecResult execResult = core.ProgBuilder.CreateInstrOpenExcelParamConst("file", fileName);

        //--Create: D.Cell > 50 
        InstrCompColCellVal instrCompIf = core.ProgBuilder.CreateInstrCompCellVal(3, ValCompOperator.GreaterThan, 50);

        //--Create: D.Cell= 12
        InstrSetCellVal instrSetValThen = core.ProgBuilder.CreateInstrSetCellVal(3, 12);

        // If D.Cell > 50 Then D.Cell= 12
        InstrIfColThen instrIfColThen;
        execResult = core.ProgBuilder.CreateInstrIfColThen(instrCompIf, instrSetValThen, out instrIfColThen);
        Assert.IsTrue(execResult.Result);

        //--create the instr ForEach x If x Then
        execResult = core.ProgBuilder.CreateInstrOnExcelForEachRowIfThen("file", 0, 1, instrIfColThen);

        Assert.IsNotNull(execResult);
        Assert.IsTrue(execResult.Result);
    }

    [TestMethod]
    public void CreateInstrForEachCellInColsIfThenNoOpenExcelError()
    {
        LexerowCore core = new LexerowCore();

        //--Create: D.Cell > 50 
        InstrCompColCellVal instrCompIf = core.ProgBuilder.CreateInstrCompCellVal(3, ValCompOperator.GreaterThan, 50);

        //--Create: D.Cell= 12
        InstrSetCellVal instrSetValThen = core.ProgBuilder.CreateInstrSetCellVal(3, 12);

        // If D.Cell > 50 Then D.Cell= 12
        InstrIfColThen instrIfColThen;
        ExecResult execResult = core.ProgBuilder.CreateInstrIfColThen(instrCompIf, instrSetValThen, out instrIfColThen);
        Assert.IsTrue(execResult.Result);

        //--create the instr ForEach x If x Then
        execResult = core.ProgBuilder.CreateInstrOnExcelForEachRowIfThen("file", 0, 1, instrIfColThen);

        Assert.IsNotNull(execResult);
        Assert.IsFalse(execResult.Result);
        Assert.AreEqual(ErrorCode.ExcelFileObjectNameDoesNotExists, execResult.ListError[0].ErrorCode);
    }

    // param execObjectFileName is null, does not defined before
    [TestMethod]
    public void CreateInstrForEachCellInColsIfThenExcelObjNullError()
    {
        LexerowCore core = new LexerowCore();

        //--Create: D.Cell > 50 
        InstrCompColCellVal instrCompIf = core.ProgBuilder.CreateInstrCompCellVal(3, ValCompOperator.GreaterThan, 50);

        //--Create: D.Cell= 12
        InstrSetCellVal instrSetValThen = core.ProgBuilder.CreateInstrSetCellVal(3, 12);

        // If D.Cell > 50 Then D.Cell= 12
        InstrIfColThen instrIfColThen;
        ExecResult execResult = core.ProgBuilder.CreateInstrIfColThen(instrCompIf, instrSetValThen, out instrIfColThen);
        Assert.IsTrue(execResult.Result);

        //--create the instr ForEach x If x Then
        execResult = core.ProgBuilder.CreateInstrOnExcelForEachRowIfThen(null, 0, 1, instrIfColThen);

        Assert.IsNotNull(execResult);
        Assert.IsFalse(execResult.Result);
        Assert.AreEqual(ErrorCode.ExcelFileObjectNameIsNull, execResult.ListError[0].ErrorCode);
    }

    // instrIfCond type should be InstrCompCellVal
}
