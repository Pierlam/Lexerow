using Lexerow.Core.System.ScriptDef;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;
public class InstrFirstRow : InstrBase
{
    public InstrFirstRow(ScriptToken scriptToken) : base(scriptToken)
    {
        InstrType = InstrType.FirstRow;
    }
}
