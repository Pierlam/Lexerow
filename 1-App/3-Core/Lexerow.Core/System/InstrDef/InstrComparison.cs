using Lexerow.Core.System.ScriptDef;

namespace Lexerow.Core.System.InstrDef;

/// <summary>
/// Instruction bool comparison.
/// Has 2 operands which are both mandatory.
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

    /// <summary>
    /// The left operand.
    /// </summary>
    public InstrBase OperandLeft { get; set; } = null;

    /// <summary>
    /// The right operand.
    /// </summary>
    public InstrBase OperandRight { get; set; } = null;

    /// <summary>
    /// The comparison operator.
    /// </summary>
    public InstrSepComparison Operator { get; set; } = null;

    /// <summary>
    /// Used during execution of the instruction.
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