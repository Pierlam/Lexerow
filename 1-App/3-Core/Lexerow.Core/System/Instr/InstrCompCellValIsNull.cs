using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;

/// <summary>
/// Comparison instruction:
/// A.Cell = null  or <>
/// </summary>
public class InstrCompCellValIsNull : InstrBase
{
    public InstrCompCellValIsNull(int colNum, InstrCompCellValOperator oper)
    {
        InstrType = InstrType.CompCellValIsNull;
        ColNum = colNum;
        Operator = oper;
    }

    public int ColNum { get; set; }

    public InstrCompCellValOperator Operator { get; set; }

}
