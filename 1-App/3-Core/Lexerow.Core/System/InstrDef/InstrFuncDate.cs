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
        IsFunctionCall = true;
        // return a date
        ReturnType = InstrFunctionReturnType.ValueDate;
    }

    public InstrBase InstrYear { get; set; }

    public InstrBase InstrMonth { get; set; }

    public InstrBase InstrDay { get; set; }

    public override string ToString()
    {
        string year = "(null)";
        if(InstrYear!=null)year= InstrYear.ToString();
        string month = "(null)";
        if (InstrMonth != null) month = InstrMonth.ToString();
        string day = "(null)";
        if (InstrDay != null) day = InstrDay.ToString();

        return "Date( " + year +"," + month +"," + month+")";
    }

}
