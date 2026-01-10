using Lexerow.Core.System.ScriptDef;

namespace Lexerow.Core.System.InstrDef;

public class InstrEndOnExcel : InstrBase
{
    public InstrEndOnExcel(ScriptToken scriptToken) : base(scriptToken)
    {
        InstrType = InstrType.EndOnExcel;
    }
}