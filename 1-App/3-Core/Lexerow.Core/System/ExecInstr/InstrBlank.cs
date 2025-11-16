using Lexerow.Core.System.ScriptDef;

namespace Lexerow.Core.System;

/// <summary>
/// 2 use cases.
/// If A.Cell=blank
/// Then A.Cell=Blank
/// </summary>
public class InstrBlank : InstrBase
{
    public InstrBlank(ScriptToken scriptToken) : base(scriptToken)
    {
        InstrType = InstrType.InstrBlank;
    }
}