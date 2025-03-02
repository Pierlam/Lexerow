using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;

public enum InstrType
{
    OpenExcel,
    CloseExcel,
    
    SetCellVal,
    SetCellNull,
    SetCellBlank,

    CompCellVal,
    CompCellValIsNull,

    InstrForEachCellInColsIfThen,
}

/// <summary>
/// Base of all instruction.
/// </summary>
public abstract class InstrBase
{
    public InstrType InstrType { get; set; }
}
