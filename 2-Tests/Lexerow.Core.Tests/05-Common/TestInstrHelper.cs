using Lexerow.Core.System;

namespace Lexerow.Core.Tests._05_Common;

/// <summary>
/// Instruction test helper.
/// </summary>
public class TestInstrHelper
{
    public static void TestInstrColCellFuncValue(string stage, InstrBase instrBase, string colName, int colNum)
    {
        InstrColCellFunc instrColCellFunc = instrBase as InstrColCellFunc;
        Assert.IsNotNull(instrColCellFunc);

        Assert.AreEqual(colName, instrColCellFunc.ColName);
        Assert.AreEqual(colNum, instrColCellFunc.ColNum);
        Assert.AreEqual(InstrColCellFuncType.Value, instrColCellFunc.InstrColCellFuncType);
    }

    public static void TestInstrConstValue(string stage, InstrBase instrBase, int val)
    {
        // check If-Operand Right
        InstrConstValue instrConstValue = instrBase as InstrConstValue;
        Assert.IsNotNull(instrConstValue);
        Assert.AreEqual(val.ToString(), instrConstValue.RawValue);
        Assert.AreEqual(val, (instrConstValue.ValueBase as ValueInt).Val);
    }
}