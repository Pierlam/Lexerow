using Lexerow.Core.System.ScriptDef;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;
public class InstrRow : InstrBase
{
    public InstrRow(ScriptToken scriptToken) : base(scriptToken)
    {
        InstrType = InstrType.Row;
    }
}
