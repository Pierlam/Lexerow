using Lexerow.Core.System.Compilator;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;
public class InstrEndOnExcel : InstrBase
{
    public InstrEndOnExcel(ScriptToken scriptToken) : base(scriptToken)
    {
        InstrType = InstrType.EndOnExcel;
    }
}
