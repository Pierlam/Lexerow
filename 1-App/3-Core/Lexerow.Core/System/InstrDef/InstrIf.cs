using Lexerow.Core.System.ScriptDef;

namespace Lexerow.Core.System.InstrDef;

/// <summary>
/// 2 cases:
/// 1/ If LeftOperand SepComparison RightOperand Then
/// 2/ If Operand Then
/// </summary>
public class InstrIf : InstrBase
{
    public InstrIf(ScriptToken scriptToken) : base(scriptToken)
    {
        InstrType = InstrType.If;
        ReturnType = InstrFunctionReturnType.ValueBool;
    }

    /// <summary>
    /// instruction which return a bool value.
    /// </summary>
    public InstrBase InstrBase { get; set; } = null;

    /// <summary>
    /// Result of the execution of the if instruction.
    /// </summary>
    public bool Result { get; set; } = false;

    public override string ToString()
    {
        return "If: ";
    }

}