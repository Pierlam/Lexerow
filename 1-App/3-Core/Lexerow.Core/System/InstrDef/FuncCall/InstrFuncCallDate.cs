using Lexerow.Core.System.ScriptDef;

namespace Lexerow.Core.System.InstrDef.FuncCall;

public class InstrFuncCallDate : InstrBase
{
    public InstrFuncCallDate(ScriptToken scriptToken) : base(scriptToken)
    {
        InstrType = InstrType.FuncDate;
        IsFunctionCall = true;
        // return a date
        ReturnType = InstrReturnType.ValueDateOnly;
    }

    public InstrBase InstrYear { get; set; }

    public InstrBase InstrMonth { get; set; }

    public InstrBase InstrDay { get; set; }

    public override string ToString()
    {
        string year = "(null)";
        if (InstrYear != null) year = InstrYear.ToString();
        string month = "(null)";
        if (InstrMonth != null) month = InstrMonth.ToString();
        string day = "(null)";
        if (InstrDay != null) day = InstrDay.ToString();

        return "FuncDate( " + year + "," + month + "," + day + ")";
    }
}