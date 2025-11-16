using Lexerow.Core.System.ScriptDef;

namespace Lexerow.Core.System;

public class InstrCol : InstrBase
{
    public InstrCol(ScriptToken scriptToken) : base(scriptToken)
    {
        InstrType = InstrType.Col;
    }
}