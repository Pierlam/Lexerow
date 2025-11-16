using Lexerow.Core.System.ScriptDef;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;
internal class InstrCharMul :InstrBase
{
    public InstrCharMul(ScriptToken scriptToken) : base(scriptToken)
    {
        InstrType = InstrType.Plus;
    }
}
