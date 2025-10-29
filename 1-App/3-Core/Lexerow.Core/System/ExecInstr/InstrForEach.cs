using Lexerow.Core.System.Compilator;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;

/// <summary>
/// Instr ForEach. TO manage the token found in the script.
/// </summary>
public class InstrForEach : InstrBase
{
    public InstrForEach(ScriptToken scriptToken) : base(scriptToken)
    {
        InstrType = InstrType.ForEach;
    }


    /// <summary>
    /// list of instr to execute. 
    /// If Then
    /// or direct instr to execute for each row.
    /// </summary>
    //public List<InstrBase> ListInstr { get; set; } = new List<InstrBase>();

}
