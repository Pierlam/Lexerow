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
public class InstrIfThenElse : InstrBase
{
    public InstrIfThenElse(ScriptToken scriptToken):base(scriptToken)
    {
        InstrType = InstrType.IfThenElse;
    }

    /// <summary>
    /// If -comparison- 
    /// can be a comparison or an fct call, should return a bool.
    /// </summary>
    public InstrBase InstrIf { get; set; }

    /// <summary>
    /// If -comparison- Then InstrThen  Else InstrElse
    /// </summary>
    public InstrBase InstrThen { get; set; }

    public InstrBase InstrElse { get; set; }

}
