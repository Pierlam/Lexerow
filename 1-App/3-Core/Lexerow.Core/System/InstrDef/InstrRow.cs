using Lexerow.Core.System.ScriptDef;

namespace Lexerow.Core.System.InstrDef;

public class InstrRow : InstrBase
{
    public InstrRow(ScriptToken scriptToken) : base(scriptToken)
    {
        InstrType = InstrType.Row;
    }
}