using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;


public enum InstrCompCellValOperator
{
    Equal,
    NotEqual,
    GreaterThan,
    LesserThan,
    GreaterOrEqualThan,
    LesserOrEqualThan,

    // valMin < val < valMax   -> InstrInIntervalCellVal
    //Between,
    //BetweenEqual,

    //NotBetween,
    ///NotBetweenEqual,
}

/// <summary>
/// Instruction comparison for cell value.
/// format: cellVal operator operand
/// exp: cellVal > 12
/// </summary>
public class InstrCompCellVal : InstrBase
{
    public InstrCompCellVal(int colNum, InstrCompCellValOperator oper, ValueBase value)
    {
        InstrType = InstrType.CompCellVal;
        ColNum = colNum;
        Operator = oper;
        Value = value;
    }

    public int ColNum { get; set; }

    public InstrCompCellValOperator Operator { get; set; }

    // operand, always a value?
    public ValueBase Value { get; set; }
}
