using Lexerow.Core.System.ScriptDef;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;

/// <summary>
/// Used to match special case in script: ForEachRow.
/// Same as ForEach Row. 
/// </summary>
public class InstrForEachRow : InstrBase
{
    public InstrForEachRow(ScriptToken scriptToken) : base(scriptToken)
    {
        InstrType = InstrType.ForEachRow;
    }
}
