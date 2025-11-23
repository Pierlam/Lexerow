using Lexerow.Core.System.ScriptDef;

namespace Lexerow.Core.System.InstrDef;
public class InstrFirstRow : InstrBase
{
    public InstrFirstRow(ScriptToken scriptToken) : base(scriptToken)
    {
        InstrType = InstrType.FirstRow;
    }
}
