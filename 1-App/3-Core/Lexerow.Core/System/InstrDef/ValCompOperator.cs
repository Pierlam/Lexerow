namespace Lexerow.Core.System.InstrDef;

/// <summary>
/// Value comparison operator.
/// </summary>
public enum ValCompOperator
{
    Equal,
    NotEqual,
    GreaterThan,
    LessThan,
    GreaterOrEqualThan,
    LessOrEqualThan,
}

/// <summary>
/// Interval comparison operator.
/// </summary>
public enum IntervalCompOperator
{
    // valMin < val < valMax
    Between,

    // valMin <= val <= valMax
    BetweenEqual,

    // valMin <= val < valMax
    BetweenEqualMin,

    // valMin < val <= valMax
    BetweenEqualMax,

    NotBetween,

    NotBetweenEqual,

    NotBetweenEqualMin,

    NotBetweenEqualMax,
}