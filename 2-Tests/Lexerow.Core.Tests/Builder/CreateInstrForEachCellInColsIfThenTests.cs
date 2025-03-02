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

        ExecResult execResult = core.Builder.CreateInstrOpenExcel("file", fileName);

        //--Create: col D, c.Value > 50 
        InstrCompCellVal instrIfCond = core.Builder.CreateInstrCompCellVal(3, InstrCompCellValOperator.GreaterThan, 50);

        //--Create: colD, c.Value:=12
        InstrSetCellVal instrSetVal = core.Builder.CreateInstrSetCellVal(3, 12);
        List<InstrBase> listInstrThen = new List<InstrBase>();
        listInstrThen.Add(instrSetVal);

        //--create the instr ForEach x If x Then
        execResult = core.Builder.CreateInstrForEachRowIfThen("file", 0, 1, instrIfCond, listInstrThen);

        Assert.IsNotNull(execResult);
        Assert.IsTrue(execResult.Result);
    }

    [TestMethod]
    public void CreateInstrForEachCellInColsIfThenNoOpenExcelError()
    {
        LexerowCore core = new LexerowCore();

        //--Create: col D, c.Value > 50 
        InstrCompCellVal instrIfCond = core.Builder.CreateInstrCompCellVal(3, InstrCompCellValOperator.GreaterThan, 50);

        //--Create: colD, c.Value:=12
        InstrSetCellVal instrSetVal = core.Builder.CreateInstrSetCellVal(3, 12);
        List<InstrBase> listInstrThen = new List<InstrBase>();
        listInstrThen.Add(instrSetVal);

        //--create the instr ForEach x If x Then
        ExecResult execResult = core.Builder.CreateInstrForEachRowIfThen("file", 0, 1, instrIfCond, listInstrThen);

        Assert.IsNotNull(execResult);
        Assert.IsFalse(execResult.Result);
        Assert.AreEqual(ErrorCode.ExcelFileObjectNameDoesNotExists, execResult.ListError[0].ErrorCode);
    }

    // param execObjectFileName is null, does not defined before
    [TestMethod]
    public void CreateInstrForEachCellInColsIfThenExcelObjNullError()
    {
        LexerowCore core = new LexerowCore();

        //--Create: col D, c.Value > 50 
        InstrCompCellVal instrIfCond = core.Builder.CreateInstrCompCellVal(3, InstrCompCellValOperator.GreaterThan, 50);

        //--Create: colD, c.Value:=12
        InstrSetCellVal instrSetVal = core.Builder.CreateInstrSetCellVal(3, 12);
        List<InstrBase> listInstrThen = new List<InstrBase>();
        listInstrThen.Add(instrSetVal);

        //--create the instr ForEach x If x Then
        ExecResult execResult = core.Builder.CreateInstrForEachRowIfThen(null, 0, 1, instrIfCond, listInstrThen);

        Assert.IsNotNull(execResult);
        Assert.IsFalse(execResult.Result);
        Assert.AreEqual(ErrorCode.ExcelFileObjectNameIsNull, execResult.ListError[0].ErrorCode);
    }

    // instrIfCond type should be InstrCompCellVal
}
