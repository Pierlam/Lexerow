using Lexerow.Core.System.ScriptDef;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
