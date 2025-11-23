namespace Lexerow.Core.System;

/// <summary>
/// manage only basic type.
/// no dateTime, will be in a string.
/// </summary>
public enum ValueType
{
    Bool,

    String,
    Int,
    Double,
    DateOnly,
    DateTime,
    TimeOnly,

    /// <summary>
    /// Token or keyword, will be a variable name defined before used.
    /// </summary>
    //Token,

    ListOfString,

    ListOfInt,
    ListOfDouble,
}

public abstract class ValueBase
{
    public ValueType ValueType { get; set; } = ValueType.String;
}

public class ValueString : ValueBase
{
    public ValueString(string val)
    {
        ValueType = ValueType.String;
        Val = val;
    }

    /// <summary>
    /// The string, without the double quote.
    /// </summary>
    public string Val { get; set; }

    /// <summary>
    /// max length of the string.
    /// </summary>
    public int MaxLength { get; set; } = int.MaxValue;
}

/// <summary>
/// Integer value.
/// </summary>
public class ValueInt : ValueBase
{
    public ValueInt(int val)
    {
        ValueType = ValueType.Int;
        Val = val;
    }

    public int Val { get; set; }
}

/// <summary>
/// Double value.
/// </summary>
public class ValueDouble : ValueBase
{
    public ValueDouble(double val)
    {
        ValueType = ValueType.Double;
        Val = val;
    }

    public double Val { get; set; }
}

/// <summary>
/// Boolean value.
/// </summary>
public class ValueBool : ValueBase
{
    public ValueBool(bool val)
    {
        ValueType = ValueType.Bool;
        Val = val;
    }

    public bool Val { get; set; }
}

/// <summary>
/// DateTime value.
/// </summary>
public class ValueDateTime : ValueBase
{
    public ValueDateTime(DateTime val)
    {
        ValueType = ValueType.DateTime;
        Val = val;
    }

    public double ToDouble()
    {
        return Val.ToOADate();
    }

    public DateTime Val { get; set; }
}

/// <summary>
/// DateOnly value.
/// </summary>
public class ValueDateOnly : ValueBase
{
    public ValueDateOnly(DateOnly val)
    {
        ValueType = ValueType.DateOnly;
        Val = val;
    }

    public DateOnly Val { get; set; }

    public DateTime ToDateTime()
    {
        return Val.ToDateTime(TimeOnly.MinValue);
    }

    public double ToDouble()
    {
        DateTime val = Val.ToDateTime(TimeOnly.MinValue);

        return val.ToOADate();
    }
}

/// <summary>
/// Time value.
/// </summary>
public class ValueTimeOnly : ValueBase
{
    public ValueTimeOnly(TimeOnly val)
    {
        ValueType = ValueType.TimeOnly;
        Val = val;
    }

    public TimeOnly Val { get; set; }

    public double ToDouble()
    {
        // set the hour, minute, second and millisecond
        DateTime dtVal = new DateTime(2025, 1, 1, Val.Hour, Val.Minute, Val.Second, Val.Millisecond);

        double res = dtVal.ToOADate();
        return res = res - Math.Truncate(res);
    }
}

public class ValueListOfString : ValueBase
{
    public ValueListOfString()
    {
        ValueType = ValueType.ListOfString;
    }

    public ValueListOfString(List<string> listVal)
    {
        ValueType = ValueType.ListOfString;
        ListVal.AddRange(listVal);
    }

    public List<string> ListVal { get; private set; } = new List<string>();
}

public class ValueListOfInt : ValueBase
{
    public ValueListOfInt(List<int> listVal)
    {
        ValueType = ValueType.ListOfInt;
        ListVal.AddRange(listVal);
    }

    public List<int> ListVal { get; private set; } = new List<int>();
}

public class ValueListOfDouble : ValueBase
{
    public ValueListOfDouble(List<double> listVal)
    {
        ValueType = ValueType.ListOfDouble;
        ListVal.AddRange(listVal);
    }

    public List<double> ListVal { get; private set; } = new List<double>();
}