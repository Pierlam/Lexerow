using Lexerow.Core.System.ScriptDef;

namespace Lexerow.Core.System.InstrDef;

public class InstrDot : InstrBase
{
    public InstrDot(ScriptToken scriptToken) : base(scriptToken)
    {
        InstrType = InstrType.Dot;
    }

    public override string ToString()
    {
        return ".";
    }

}