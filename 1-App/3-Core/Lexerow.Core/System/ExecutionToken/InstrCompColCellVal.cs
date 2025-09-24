using Lexerow.Core.System.Compilator;
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
    public InstrCompColCellVal(ScriptToken scriptToken, int colNum, ValCompOperator oper, ValueBase value):base(scriptToken)
    {
        InstrType = InstrType.CompCellVal;
        ColNum = colNum;
        Operator = oper;
        Value = value;
    }

    public int ColNum { get; set; }

    public ValCompOperator Operator { get; set; }

    public ValueBase Value { get; set; }
}
