using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;

/// <summary>
/// Delete a cell.
/// Remove the value and the style.
/// </summary>
public class InstrSetCellNull : InstrBase
{
    public InstrSetCellNull(int colNum)
    {
        InstrType = InstrType.SetCellNull;
        ColNum = colNum;
    }

    /// <summary>
    /// column number.
    /// Shoud be > 0
    /// </summary>
    public int ColNum { get; set; }
}
