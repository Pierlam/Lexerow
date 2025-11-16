using Lexerow.Core.System.ScriptDef;

namespace Lexerow.Core.System;

public class InstrNull : InstrBase
{
    public InstrNull(ScriptToken scriptToken) : base(scriptToken)
    {
        InstrType = InstrType.InstrNull;
    }
}