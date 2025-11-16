using Lexerow.Core.System.ScriptDef;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;

/// <summary>
/// Process a data row. Execute all defined instructions.
/// ForEach Row; IfThen...
/// Used during execution.
/// </summary>
public class InstrProcessRow : InstrBase
{
    public InstrProcessRow(ScriptToken scriptToken, List<InstrBase> listInstrForEachRow) : base(scriptToken)
    {
        InstrType = InstrType.ProcessRow;
        ListInstrForEachRow = listInstrForEachRow;
    }

    /// <summary>
    /// 0-based.
    /// </summary>
    public int RowNum { get; set; } = -1;

    public int ColNum { get; set; } = -1;   

    /// <summary>
    /// List of Instructions to execute on each datarow.
    /// </summary>
    public List<InstrBase> ListInstrForEachRow { get; set; }
}
