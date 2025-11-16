using Lexerow.Core.System.ScriptDef;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;

/// <summary>
/// Instruction process instruction for each datarow.
/// Instr are defined in ForEach Row block.
/// Used during execution.
/// </summary>
public class InstrProcessInstrForEachRow : InstrBase
{
    public InstrProcessInstrForEachRow(ScriptToken scriptToken, List<InstrBase> listInstr) : base(scriptToken)
    {
        InstrType = InstrType.ProcessInstrForEachRow;
        ListInstr.AddRange(listInstr);
    }

    public int InstrToProcessNum { get; set; } = -1;

    /// <summary>
    /// OnSheet/ForEach Row, list of instruction to execute.
    /// </summary>
    public List<InstrBase> ListInstr { get; set; } = new List<InstrBase>();
}
