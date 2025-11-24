using Lexerow.Core.System;
using Lexerow.Core.System.InstrDef;

namespace Lexerow.Core.Tests._05_Common;

/// <summary>
/// Instruction test helper.
/// </summary>
public class TestInstrHelper
{
    public static int GetValueInt(InstrBase instr)
    {
        return ((instr as InstrValue).ValueBase as ValueInt).Val;
    }

    public static void TestInstrColCellFuncValue(string stage, InstrBase instrBase, string colName, int colNum)
    {
        InstrColCellFunc instrColCellFunc = instrBase as InstrColCellFunc;
        Assert.IsNotNull(instrColCellFunc);

        Assert.AreEqual(colName, instrColCellFunc.ColName);
        Assert.AreEqual(colNum, instrColCellFunc.ColNum);
        Assert.AreEqual(InstrColCellFuncType.Value, instrColCellFunc.InstrColCellFuncType);
    }

    public static void TestInstrValue(string stage, InstrBase instrBase, int val)
    {
        // check If-Operand Right
        InstrValue instrValue = instrBase as InstrValue;
        Assert.IsNotNull(instrValue);
        Assert.AreEqual(val.ToString(), instrValue.RawValue);
        Assert.AreEqual(val, (instrValue.ValueBase as ValueInt).Val);
    }
}