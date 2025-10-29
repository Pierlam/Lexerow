using Lexerow.Core.System.Compilator;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;

/// <summary>
/// Used during execution.
/// </summary>
public class InstrNextRow : InstrBase
{
    public InstrNextRow(ScriptToken scriptToken, List<InstrBase> listInstrForEachRow) : base(scriptToken)
    {
        InstrType = InstrType.NextRow;
        ListInstrForEachRow = listInstrForEachRow;
    }

    /// <summary>
    /// 0-based.
    /// </summary>
    public int RowNum { get; set; } = -1;

    public int ColNum { get; set; } = -1;   
    public List<InstrBase> ListInstrForEachRow { get; set; }
}
