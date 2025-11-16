using Lexerow.Core.System.ScriptDef;

namespace Lexerow.Core.System;

public class InstrOpenBracket : InstrBase
{
    public InstrOpenBracket(ScriptToken scriptToken) : base(scriptToken)
    {
        InstrType = InstrType.OpenBracket;
    }
}