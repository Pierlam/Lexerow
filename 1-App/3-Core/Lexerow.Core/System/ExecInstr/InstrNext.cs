using Lexerow.Core.System.ScriptDef;

namespace Lexerow.Core.System;

/// <summary>
///  ForEach Row
///     instr
///   Next
/// </summary>
public class InstrNext : InstrBase
{
    public InstrNext(ScriptToken scriptToken) : base(scriptToken)
    {
        InstrType = InstrType.Next;
    }
}