using Lexerow.Core.System.ScriptDef;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;
internal class InstrCharPlus : InstrBase
{
    public InstrCharPlus(ScriptToken scriptToken) : base(scriptToken)
    {
        InstrType = InstrType.Plus;
    }
}
