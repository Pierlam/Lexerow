using Lexerow.Core.System.ScriptDef;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;
public class InstrThen:InstrBase
{
    public InstrThen(ScriptToken scriptToken):base(scriptToken)
    {
        InstrType = InstrType.Then;
    }


    /// <summary>
    /// 1/ There is an instr on the script line than the token Then.
    /// exp: Then instr
    /// The EndIf instr is not expected.
    ///
    /// 2/ There is NO instr on the script line than the token Then.
    /// exp:  Then
    ///          Instr
    ///       End If
    /// So in this case, the token End If Is expected.      
    /// </summary>
    //public bool IsEndIfInstrExpected { get; set; } = false;

    public bool HasInstrAfterInSameLine { get; set; } = false;

    /// <summary>
    /// Used in parse process.
    /// </summary>
    public bool IsEndIfReached { get; set; } = false;

    /// <summary>
    /// list of instr to execute. 
    /// </summary>
    public List<InstrBase> ListInstr { get;set; }=new List<InstrBase>();

    /// <summary>
    /// then current instr to execute.
    /// </summary>
    public int RunInstrNum { get; set; } = -1;


    public void ClearRun()
    {
        RunInstrNum = -1;
    }
}
