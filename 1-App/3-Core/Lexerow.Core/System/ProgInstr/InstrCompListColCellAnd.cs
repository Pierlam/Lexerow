using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;

/// <summary>
/// Instruction compare a list column value with AND between each of us.
/// </summary>
public class InstrCompListColCellAnd : InstrRetBoolBase
{
    public InstrCompListColCellAnd(List<InstrRetBoolBase> listInstrCompIf)
    {
        InstrType = InstrType.CompListColCellAnd;
        ListInstrComp.AddRange(listInstrCompIf);
    }

    public List<InstrRetBoolBase> ListInstrComp { get; private set; } = new List<InstrRetBoolBase>();
}
