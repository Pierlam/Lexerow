using Lexerow.Core.System.Compilator;
using NPOI.XSSF.Streaming.Values;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;

/// <summary>
/// Represents a const value.
/// can be an int, double, string, ...
/// </summary>
public class InstrConstValue : InstrBase
{
    public InstrConstValue(ScriptToken scriptToken, string rawValue): base(scriptToken)
    {
        InstrType = InstrType.ConstValue;
        RawValue = rawValue;
        ValueBase= new ValueString(rawValue);
    }

    public InstrConstValue(ScriptToken scriptToken, int value) : base(scriptToken)
    {
        InstrType = InstrType.ConstValue;
        RawValue = value.ToString();
        ValueBase = new ValueInt(value);
    }

    public InstrConstValue(ScriptToken scriptToken, double value) : base(scriptToken)
    {
        InstrType = InstrType.ConstValue;
        RawValue = value.ToString();
        ValueBase = new ValueDouble(value);
    }

    public string RawValue { get; set; }

    public ValueBase ValueBase { get; set; }

    public override string ToString()
    {
        return "Type: ConstValue, Val: " + RawValue;
    }
}
