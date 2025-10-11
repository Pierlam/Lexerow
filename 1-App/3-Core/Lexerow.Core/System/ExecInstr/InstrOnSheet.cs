using Lexerow.Core.System.Compilator;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;

/// <summary>
/// OnSheet SheetNum/SheetName
///    [FirstRow <val>] [FirstCol <val>]
///    	[OnHeader]  todo: To Be Defined
///    ForEach Row
///         instr..
/// </summary>
public class InstrOnSheet : InstrBase
{
    public InstrOnSheet(ScriptToken scriptToken) : base(scriptToken)
    {
        InstrType = InstrType.OnSheet;
    }

    public int SheetNum { get; set; } = 1;

    public string SheetName { get; set; }=string.Empty;

    public int FirstRowNum { get; set; } = 1;

    public int FirstColNum { get; set; } = 1;

    /// <summary>
    /// List of instr to apply on each row.
    /// ForEach Row
    /// </summary>
    public List<InstrBase> ListInstr { get; set; }= new List<InstrBase>();  
}
