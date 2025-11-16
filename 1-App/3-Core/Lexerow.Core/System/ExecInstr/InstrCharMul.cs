using Lexerow.Core.System.ScriptDef;

namespace Lexerow.Core.System;

internal class InstrCharMul : InstrBase
{
    public InstrCharMul(ScriptToken scriptToken) : base(scriptToken)
    {
        InstrType = InstrType.Plus;
    }
}