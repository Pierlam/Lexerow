using Lexerow.Core.System.ScriptDef;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System.InstrDef;

internal class InstrOr : InstrBase
{
    public InstrOr(ScriptToken scriptToken) : base(scriptToken)
    {
        InstrType = InstrType.Or;
    }

    public override string ToString()
    {
        return "Or";
    }

}
