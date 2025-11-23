using Lexerow.Core.System.ScriptDef;

namespace Lexerow.Core.System.InstrDef;

internal class InstrCharMul : InstrBase
{
    public InstrCharMul(ScriptToken scriptToken) : base(scriptToken)
    {
        InstrType = InstrType.Plus;
    }
}