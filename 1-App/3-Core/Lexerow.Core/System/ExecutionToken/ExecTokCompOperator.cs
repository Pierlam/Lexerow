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

public class ExecTokCompOperator : ExecTokBase
{
    /// <summary>
    /// TODO: keep it?
    /// </summary>
    /// <param name="value"></param>
    /// <param name="type"></param>
    public ExecTokCompOperator(string value, ExecTokCompOperatorType type): base (null)
    {
        ExecTokType = ExecTokType.CompOperator;
        Value = value;
        Type = type;
    }

    public string Value { get; set; }   

    public ExecTokCompOperatorType Type { get; set; }   
}
