using Lexerow.Core.System.Compilator;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;
public class InstrForEachRow : InstrBase
{
    public InstrForEachRow(ScriptToken scriptToken, List<InstrBase> listInstr) : base(scriptToken)
    {
        InstrType = InstrType.ForEachRow;
        ListInstr.AddRange(listInstr);
    }

    public int InstrToProcessNum { get; set; } = -1;

    /// <summary>
    /// OnSheet/ForEach Row, list of instruction to execute.
    /// </summary>
    public List<InstrBase> ListInstr { get; set; } = new List<InstrBase>();
}
