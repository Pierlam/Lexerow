using Lexerow.Core.System.ScriptDef;

namespace Lexerow.Core.System;

public class InstrEndIf : InstrBase
{
    public InstrEndIf(ScriptToken scriptToken) : base(scriptToken)
    {
        InstrType = InstrType.EndIf;
    }
}