using Lexerow.Core.System.InstrDef;
using Lexerow.Core.System.ScriptDef;

namespace Lexerow.Core.System;

public class InstrEndIf : InstrBase
{
    public InstrEndIf(ScriptToken scriptToken) : base(scriptToken)
    {
        InstrType = InstrType.EndIf;
    }

    public override string ToString()
    {
        return "EndIf: ";
    }

}