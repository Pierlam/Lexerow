using Lexerow.Core.System.Compilator;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;
public class InstrEnd : InstrBase
{
    public InstrEnd(ScriptToken scriptToken) : base(scriptToken)
    {
        InstrType = InstrType.End;
    }
}
