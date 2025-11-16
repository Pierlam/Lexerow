using Lexerow.Core.System.ScriptDef;

namespace Lexerow.Core.System;

public class InstrComma : InstrBase
{
    public InstrComma(ScriptToken scriptToken) : base(scriptToken)
    {
        InstrType = InstrType.Comma;
    }
}