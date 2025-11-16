using Lexerow.Core.System.ScriptDef;

namespace Lexerow.Core.System;

internal class InstrCharDiv : InstrBase
{
    public InstrCharDiv(ScriptToken scriptToken) : base(scriptToken)
    {
        InstrType = InstrType.Plus;
    }
}