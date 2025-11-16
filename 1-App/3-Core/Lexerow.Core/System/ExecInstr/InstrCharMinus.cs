using Lexerow.Core.System.ScriptDef;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;
internal class InstrCharMinus:InstrBase
{
    public InstrCharMinus(ScriptToken scriptToken) : base(scriptToken)
    {
        InstrType = InstrType.Minus;
    }
}
