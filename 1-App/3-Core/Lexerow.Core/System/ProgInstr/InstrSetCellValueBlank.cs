using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;

/// <summary>
/// Clean the value of the cell.
/// Does not delete the cell.
/// Remove only the value, keep the style.
/// </summary>
public class InstrSetCellValueBlank : InstrBase
{
    public InstrSetCellValueBlank(int colNum)
    {
        InstrType = InstrType.SetCellBlank;
        ColNum = colNum;
    }

    /// <summary>
    /// column number.
    /// Shoud be > 0
    /// </summary>
    public int ColNum { get; set; }
}
