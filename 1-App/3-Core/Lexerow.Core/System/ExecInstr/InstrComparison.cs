using Lexerow.Core.System.ScriptDef;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;

/// <summary>
/// Instruction bool comparison.
/// A.Cell>10
/// 
/// </summary>
public class InstrComparison : InstrBase
{
    public InstrComparison(ScriptToken scriptToken) : base(scriptToken)
    {
        InstrType = InstrType.Comparison;
        ReturnType = InstrFunctionReturnType.ValueBool;
    }
    public InstrBase OperandLeft { get; set; } = null;
    public InstrBase OperandRight { get; set; } = null;
    public InstrSepComparison Operator { get; set; } = null;

    /// <summary>
    /// Execution, Result of the comparison.
    /// </summary>
    public bool Result { get; set; } = false;
}
