using Lexerow.Core.System.ScriptDef;

namespace Lexerow.Core.System;

internal class InstrCharPlus : InstrBase
{
    public InstrCharPlus(ScriptToken scriptToken) : base(scriptToken)
    {
        InstrType = InstrType.Plus;
    }
}