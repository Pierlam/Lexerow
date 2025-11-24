using Lexerow.Core.System.ScriptDef;

namespace Lexerow.Core.System.InstrDef;

public class InstrEnd : InstrBase
{
    public InstrEnd(ScriptToken scriptToken) : base(scriptToken)
    {
        InstrType = InstrType.End;
    }

    public override string ToString()
    {
        return "End: ";
    }

}