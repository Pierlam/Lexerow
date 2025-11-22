using Lexerow.Core.System.ScriptDef;

namespace Lexerow.Core.System;

/// <summary>
/// Represents a basic value.
/// can be an int, double, or string.
/// </summary>
public class InstrValue : InstrBase
{
    public InstrValue(ScriptToken scriptToken, string rawValue) : base(scriptToken)
    {
        InstrType = InstrType.ConstValue;
        RawValue = rawValue;
        ValueBase = new ValueString(rawValue);
    }

    public InstrValue(ScriptToken scriptToken, int value) : base(scriptToken)
    {
        InstrType = InstrType.ConstValue;
        RawValue = value.ToString();
        ValueBase = new ValueInt(value);
    }

    public InstrValue(ScriptToken scriptToken, double value) : base(scriptToken)
    {
        InstrType = InstrType.ConstValue;
        RawValue = value.ToString();
        ValueBase = new ValueDouble(value);
    }

    public string RawValue { get; set; }

    public ValueBase ValueBase { get; set; }

    public override string ToString()
    {
        return "Type: Value, Val: " + RawValue;
    }
}