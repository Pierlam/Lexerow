using Lexerow.Core.System.ScriptDef;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
