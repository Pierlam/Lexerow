using Lexerow.Core.System.ScriptDef;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;
public class InstrOpenBracket : InstrBase
{
    public InstrOpenBracket(ScriptToken scriptToken):base(scriptToken)
    {
        InstrType = InstrType.OpenBracket;
    }
}
