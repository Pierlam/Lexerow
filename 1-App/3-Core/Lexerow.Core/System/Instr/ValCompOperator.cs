using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;

/// <summary>
/// Value comparison operator.
/// </summary>
public enum ValCompOperator
{
    Equal,
    NotEqual,
    GreaterThan,
    LesserThan,
    GreaterOrEqualThan,
    LesserOrEqualThan,
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
