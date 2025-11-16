using Lexerow.Core.System.ScriptDef;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;
public class InstrEndIf : InstrBase
{
    public InstrEndIf(ScriptToken scriptToken) : base(scriptToken)
    {
        InstrType= InstrType.EndIf;
    }
}
