using Lexerow.Core.System.Compilator;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;

/// <summary>
/// 2 cases:
/// 1/ If LeftOperand SepComparison RightOperand Then
/// 2/ If Operand Then
/// </summary>
public class InstrIf:InstrBase
{
    public InstrIf(ScriptToken scriptToken):base(scriptToken)
    {
        InstrType = InstrType.If;
        ReturnType = InstrFunctionReturnType.ValueBool;
    }

    public InstrBase Operand { get; set; } = null;

    public InstrBase OperandLeft { get; set; } = null;
    public InstrBase OperandRight { get; set; } = null;
    public InstrSepComparison Operator { get; set; } = null;

}
