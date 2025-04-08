using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;


/// <summary>
/// Instruction comparison for cell value.
/// format: cellVal operator operand 
/// exp: A.Cell > 12
/// </summary>
public class InstrCompColCellVal : InstrRetBoolBase
{
    public InstrCompColCellVal(int colNum, ValCompOperator oper, ValueBase value)
    {
        InstrType = InstrType.CompCellVal;
        ColNum = colNum;
        Operator = oper;
        Value = value;
    }

    public int ColNum { get; set; }

    public ValCompOperator Operator { get; set; }

    // operand, always a value?
    public ValueBase Value { get; set; }
}
