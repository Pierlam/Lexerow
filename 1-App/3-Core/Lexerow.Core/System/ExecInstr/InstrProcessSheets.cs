using Lexerow.Core.System.Compilator;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;

/// <summary>
/// After instr OnExcel, process all sheets.
/// Used during execution.
/// </summary>
public class InstrProcessSheets : InstrBase
{
    public InstrProcessSheets(ScriptToken scriptToken, List<InstrOnSheet> listSheet) : base(scriptToken)
    {
        InstrType = InstrType.ProcessSheets;
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
