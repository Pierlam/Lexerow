using Lexerow.Core.System.ScriptDef;

namespace Lexerow.Core.System;

public class InstrEnd : InstrBase
{
    public InstrEnd(ScriptToken scriptToken) : base(scriptToken)
    {
        InstrType = InstrType.End;
    }
}