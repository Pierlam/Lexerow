using Lexerow.Core.System.Compilator;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;
internal class InstrIf:InstrBase
{
    public InstrIf(ScriptToken scriptToken):base(scriptToken)
    {
        InstrType = InstrType.If;
    }
}
