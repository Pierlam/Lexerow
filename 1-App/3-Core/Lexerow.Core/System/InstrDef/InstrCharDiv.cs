using Lexerow.Core.System.ScriptDef;

namespace Lexerow.Core.System.InstrDef;

internal class InstrCharDiv : InstrBase
{
    public InstrCharDiv(ScriptToken scriptToken) : base(scriptToken)
    {
        InstrType = InstrType.Plus;
    }
}