using Lexerow.Core.System.ScriptDef;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System.InstrDef;
public class InstrFuncDate : InstrBase
{
    public InstrFuncDate(ScriptToken scriptToken) : base(scriptToken)
    {
        InstrType = InstrType.FuncDate;
    }

    public int Year { get; set; } = 0;

    public int Month { get; set; } = 0;
    public int Day { get; set; } = 0;
}
