using Lexerow.Core.System.Compilator;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;

/// <summary>
/// If -comparison- Then InstrThen  Else InstrElse
/// </summary>
public class InstrIfThenElse : ExecTokBase
{
    public InstrIfThenElse(ScriptToken scriptToken):base(scriptToken)
    {
        ExecTokType = ExecTokType.IfThenElse;
    }

    /// <summary>
    /// If -comparison- 
    /// can be a comparison or an fct call, should return a bool.
    /// </summary>
    public ExecTokBase InstrIf { get; set; }

    /// <summary>
    /// If -comparison- Then InstrThen  Else InstrElse
    /// </summary>
    public ExecTokBase InstrThen { get; set; }

    public ExecTokBase InstrElse { get; set; }

}
