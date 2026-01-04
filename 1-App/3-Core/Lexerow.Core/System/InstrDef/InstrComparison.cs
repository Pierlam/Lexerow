using Lexerow.Core.System.ScriptDef;

namespace Lexerow.Core.System.InstrDef;

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
        ReturnType = InstrReturnType.ValueBool;
    }

    public InstrBase OperandLeft { get; set; } = null;
    public InstrBase OperandRight { get; set; } = null;
    public InstrSepComparison Operator { get; set; } = null;

    /// <summary>
    /// 0: nothing
    /// 1: Instr Left
    /// 2: Instr Right
    /// </summary>
    public int LastInstrExecuted { get; set; } = 0;

    /// <summary>
    /// Execution, Result of the comparison.
    /// </summary>
    public bool Result { get; set; } = false;

    public override string ToString()
    {
        string s = string.Empty;
        if(OperandLeft != null) s= OperandLeft.ToString();
        if (Operator != null) s += Operator.ToString();
        if (OperandRight != null) s += OperandRight.ToString();

        return "Compare " + s;
    }
}