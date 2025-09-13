using Lexerow.Core.System.Compilator;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;

/// <summary>
/// Comparison instruction:
/// A.Cell = null  or A.Cell <> null
/// </summary>
public class InstrCompColCellValIsNull : InstrRetBoolBase
{
    public InstrCompColCellValIsNull(ScriptToken scriptToken, int colNum, ValCompOperator oper):base(scriptToken)
    {
        ExecTokType = ExecTokType.CompCellValIsNull;
        ColNum = colNum;
        Operator = oper;
    }

    public int ColNum { get; set; }

    public ValCompOperator Operator { get; set; }

}
