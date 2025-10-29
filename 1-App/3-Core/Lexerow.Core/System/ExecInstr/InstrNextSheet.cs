using Lexerow.Core.System.Compilator;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;
public class InstrNextSheet : InstrBase
{
    public InstrNextSheet(ScriptToken scriptToken, List<InstrOnSheet> listSheet) : base(scriptToken)
    {
        InstrType = InstrType.NextSheet;
        ListSheet= listSheet;
    }

    /// <summary>
    /// Sheets to process.
    /// </summary>
    public List<InstrOnSheet> ListSheet { get; set; }

    /// <summary>
    /// 0-based.
    /// </summary>
    public int SheetNum { get; set; } = -1;
}
