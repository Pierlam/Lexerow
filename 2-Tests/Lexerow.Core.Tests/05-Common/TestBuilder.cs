using Lexerow.Core.System;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Tests._05_Common;
public class TestBuilder
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
