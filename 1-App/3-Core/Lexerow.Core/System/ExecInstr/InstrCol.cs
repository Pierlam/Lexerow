using Lexerow.Core.System.ScriptDef;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;
public class InstrCol : InstrBase
{
    public InstrCol(ScriptToken scriptToken) : base(scriptToken)
    {
        InstrType = InstrType.Col;
    }
}
