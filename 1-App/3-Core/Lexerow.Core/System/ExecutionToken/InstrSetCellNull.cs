using Lexerow.Core.System.Compilator;
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
public class InstrSetCellNull : ExecTokBase
{
    public InstrSetCellNull(ScriptToken scriptToken, int colNum):base(scriptToken)
    {
        ExecTokType = ExecTokType.SetCellNull;
        ColNum = colNum;
    }

    /// <summary>
    /// column number.
    /// Shoud be > 0
    /// </summary>
    public int ColNum { get; set; }
}
