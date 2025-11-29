using Lexerow.Core.System.ScriptDef;
using NPOI.HPSF;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System.InstrDef.InstrObjectDef;

/// <summary>
/// A date only object.
/// </summary>
public class InstrObjectDate : InstrBase
{
    public InstrObjectDate(ScriptToken scriptToken, DateOnly dateOnly) : base(scriptToken)
    {
        InstrType = InstrType.ObjectExcelFile;
        DateOnly = dateOnly;
    }

    public DateOnly DateOnly { get; set; }  
}
