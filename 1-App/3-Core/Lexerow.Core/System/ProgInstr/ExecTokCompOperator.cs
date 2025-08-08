using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;

public enum ExecTokCompOperatorType
{
    Equal,
    NotEqual,
    GreaterThan,
    LessThan,
    GreaterOrEqualThan,
    LessOrEqualThan,
}
public class ExecTokCompOperator : InstrBase
{
    public ExecTokCompOperator(string value, ExecTokCompOperatorType type)
    {
        InstrType = InstrType.CompOperator;
        Value = value;
        Type = type;
    }

    public string Value { get; set; }   

    public ExecTokCompOperatorType Type { get; set; }   
}
