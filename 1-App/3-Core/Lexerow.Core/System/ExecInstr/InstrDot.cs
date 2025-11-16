using Lexerow.Core.System.ScriptDef;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;
public class InstrDot : InstrBase
{
    public InstrDot(ScriptToken scriptToken) : base(scriptToken)
    {
        InstrType = InstrType.Dot;
    }
}
