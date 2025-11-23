using Lexerow.Core.System.ScriptDef;

namespace Lexerow.Core.System.InstrDef;

public class InstrCharMinus : InstrBase
{
    public InstrCharMinus(ScriptToken scriptToken) : base(scriptToken)
    {
        InstrType = InstrType.Minus;
    }
}