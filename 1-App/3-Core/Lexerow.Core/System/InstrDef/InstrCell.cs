using Lexerow.Core.System.ScriptDef;

namespace Lexerow.Core.System.InstrDef;

public class InstrCell : InstrBase
{
    public InstrCell(ScriptToken scriptToken) : base(scriptToken)
    {
        InstrType = InstrType.Cell;
    }

    public override string ToString()
    {
        return "Cell";
    }
}