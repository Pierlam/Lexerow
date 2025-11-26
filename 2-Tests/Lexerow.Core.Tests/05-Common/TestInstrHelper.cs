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

    public static bool TestInstrColCellFuncValue(InstrBase instrBase, string colName, int colNum)
    {
        InstrColCellFunc instrColCellFunc = instrBase as InstrColCellFunc;
        if(instrColCellFunc==null) return false;

        if (colName != instrColCellFunc.ColName) return false;
        if (colNum != instrColCellFunc.ColNum) return false;
        if (instrColCellFunc.InstrColCellFuncType != InstrColCellFuncType.Value) return false;

        return true;
    }

    public static bool TestInstrValue(InstrBase instrBase, int val)
    {
        // check If-Operand Right
        InstrValue instrValue = instrBase as InstrValue;
        if(instrValue==null)return false;

        if (instrValue.RawValue != val.ToString()) return false;
        if((instrValue.ValueBase as ValueInt).Val != val)return false;
        return true;
    }
}