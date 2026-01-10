using Lexerow.Core.System.ScriptDef;

namespace Lexerow.Core.System.InstrDef;

/// <summary>
/// Base of instructions returning a bool value.
/// Comparison, function
/// but not SetCell or SetVar.
/// </summary>
public class InstrRetBoolBase : InstrBase
{
    public InstrRetBoolBase(ScriptToken scriptToken) : base(scriptToken)
    {
    }
}