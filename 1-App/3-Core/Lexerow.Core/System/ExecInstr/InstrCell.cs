using Lexerow.Core.System.Compilator;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;
public class InstrCell : InstrBase
{
    public InstrCell(ScriptToken scriptToken) : base(scriptToken)
    {
        InstrType= InstrType.Cell;
    }
}
