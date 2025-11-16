using Lexerow.Core.System.ScriptDef;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;

//public enum InstrIfThenElseRunState
//{
//    None, IfInProgress, IfResultTrue, IfResultFalse, ThenInProgress, ThenAllExecuted,
//    ElseInProgress, ElseAllExecuted
//}

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
    public InstrIf InstrIf { get; set; }

    /// <summary>
    /// If -comparison- Then InstrThen  Else InstrElse
    /// </summary>
    public InstrThen InstrThen { get; set; }

    public InstrBase InstrElse { get; set; }

    // TODO: remove
    //public InstrIfThenElseRunState State { get; set; }= InstrIfThenElseRunState.None;
}
