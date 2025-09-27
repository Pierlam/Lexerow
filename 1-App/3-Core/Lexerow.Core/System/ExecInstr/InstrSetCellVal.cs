using Lexerow.Core.System.Compilator;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;

/// <summary>
/// Instruction Set Cell Value.
/// exp: Cell.Value:=12
/// </summary>
public class InstrSetCellVal :InstrBase
{
    public InstrSetCellVal(ScriptToken scriptToken, int colNum, ValueBase value):base (scriptToken)
    {
        InstrType = InstrType.SetCellVal;
        ColNum = colNum;
        Value = value;
    }

    /// <summary>
    /// column number.
    /// Shoud be > 0
    /// </summary>
    public int ColNum { get; set; } 

    /// <summary>
    /// The value
    /// </summary>
    public ValueBase Value { get; set; }
}
