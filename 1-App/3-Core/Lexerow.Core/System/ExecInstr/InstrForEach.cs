using Lexerow.Core.System.ScriptDef;

namespace Lexerow.Core.System;

/// <summary>
/// Instr ForEach. To manage the token found in the script.
/// Used by the parser.
/// </summary>
public class InstrForEach : InstrBase
{
    public InstrForEach(ScriptToken scriptToken) : base(scriptToken)
    {
        InstrType = InstrType.ForEach;
    }
}