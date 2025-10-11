using Lexerow.Core.System.Compilator;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;

public enum SepComparisonOperator
{
    Undefined,
    Equal,
    Different,
    GreaterThan,
    LessThan,
    GreaterEqualThan,
    LessEqualThan,
}

public class InstrSepComparison : InstrBase
{
    public InstrSepComparison(ScriptToken scriptToken) : base(scriptToken)
    {
        InstrType = InstrType.SepComparison;

        if(scriptToken.Value=="=") Operator= SepComparisonOperator.Equal;
        if (scriptToken.Value == "<>") Operator = SepComparisonOperator.Different;
        if (scriptToken.Value == ">") Operator = SepComparisonOperator.GreaterThan;
        if (scriptToken.Value == "<") Operator = SepComparisonOperator.LessThan;
        if (scriptToken.Value == ">=") Operator = SepComparisonOperator.GreaterEqualThan;
        if (scriptToken.Value == "<=") Operator = SepComparisonOperator.LessEqualThan;        
    }

    public SepComparisonOperator Operator { get; set; }= SepComparisonOperator.Undefined;
}
