using Lexerow.Core.System.ScriptDef;

namespace Lexerow.Core.System.InstrDef;

/// <summary>
/// Represents a basic value.
/// can be an int, double, or string.
/// </summary>
public class InstrValue : InstrBase
{
    public InstrValue(ScriptToken scriptToken, string rawValue) : base(scriptToken)
    {
        InstrType = InstrType.Value;
        RawValue = rawValue;
        ValueBase = new ValueString(rawValue);
        ValueType= ValueType.String;
    }

    public InstrValue(ScriptToken scriptToken, int value) : base(scriptToken)
    {
        InstrType = InstrType.Value;
        RawValue = value.ToString();
        ValueBase = new ValueInt(value);
        ValueType= ValueType.Int;
    }

    public InstrValue(ScriptToken scriptToken, double value) : base(scriptToken)
    {
        InstrType = InstrType.Value;
        RawValue = value.ToString();
        ValueBase = new ValueDouble(value);
        ValueType= ValueType.Double;
    }

    public InstrValue(ScriptToken scriptToken, DateTime value) : base(scriptToken)
    {
        InstrType = InstrType.Value;
        RawValue = value.ToString();
        ValueBase = new ValueDateTime(value);
        ValueType = ValueType.DateTime;
    }

    public InstrValue(ScriptToken scriptToken, DateOnly value) : base(scriptToken)
    {
        InstrType = InstrType.Value;
        RawValue = value.ToString();
        ValueBase = new ValueDateOnly(value);
        ValueType = ValueType.DateOnly;
    }

    public InstrValue(ScriptToken scriptToken, TimeOnly value) : base(scriptToken)
    {
        InstrType = InstrType.Value;
        RawValue = value.ToString();
        ValueBase = new ValueTimeOnly(value);
        ValueType = ValueType.TimeOnly;
    }

    public string RawValue { get; set; }

    /// <summary>
    /// Value object.
    /// </summary>
    public ValueBase ValueBase { get; set; }

    /// <summary>
    /// Value type.
    /// </summary>
    public ValueType ValueType { get; set; }

    public override string ToString()
    {
        return "Value: " + RawValue;
    }
}