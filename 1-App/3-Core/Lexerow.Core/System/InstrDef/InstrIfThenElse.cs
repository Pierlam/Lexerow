using Lexerow.Core.System.ScriptDef;

namespace Lexerow.Core.System.InstrDef;

/// <summary>
/// If -comparison- Then InstrThen  Else InstrElse
/// </summary>
public class InstrIfThenElse : InstrBase
{
    public InstrIfThenElse(ScriptToken scriptToken) : base(scriptToken)
    {
        InstrType = InstrType.IfThenElse;
    }

    /// <summary>
    /// If -comparison-
    /// can be a comparison or an fct call, should return a bool.
    /// </summary>
    public InstrIf InstrIf { get; set; }

    /// <summary>
    /// If -comparison- Then InstrThen  Else InstrElse
    /// </summary>
    public InstrThen InstrThen { get; set; }

    public InstrBase InstrElse { get; set; }
}