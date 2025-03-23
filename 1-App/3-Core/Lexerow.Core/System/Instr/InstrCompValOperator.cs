using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;

/// <summary>
/// InstrCompValOperator
/// </summary>
public enum InstrCompValOperator
{
    Equal,
    NotEqual,
    GreaterThan,
    LesserThan,
    GreaterOrEqualThan,
    LesserOrEqualThan,
}

public enum InstrCompIntervalValOperator
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
