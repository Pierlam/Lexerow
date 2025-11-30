using Lexerow.Core.System.ScriptDef;
using NPOI.HPSF;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System.InstrDef.Object;

/// <summary>
/// A date only object.
/// </summary>
public class InstrObjectDate : InstrBase
{
    public InstrObjectDate(ScriptToken scriptToken, ValueDateOnly valueDateOnly) : base(scriptToken)
    {
        InstrType = InstrType.ObjectExcelFile;
        //DateOnly = dateOnly;
        ValueDateOnly = valueDateOnly;
    }

    public ValueDateOnly ValueDateOnly { get; set; }

    public override string ToString()
    {
        return "ObjectDate( " + ValueDateOnly.Val.Year + "," + ValueDateOnly.Val.Month+ "," + ValueDateOnly.Val.Day + ")";
    }
}
