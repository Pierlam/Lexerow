using Lexerow.Core.System.ScriptDef;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System.InstrDef;

internal class InstrAnd : InstrBase
{
    public InstrAnd(ScriptToken scriptToken) : base(scriptToken)
    {
        InstrType = InstrType.And;
    }

    public override string ToString()
    {
        return "And";
    }
}
