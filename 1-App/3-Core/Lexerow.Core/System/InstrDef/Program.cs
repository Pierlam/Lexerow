using Lexerow.Core.System.ScriptDef;
using System.Security.AccessControl;

namespace Lexerow.Core.System.InstrDef;

/// <summary>
/// Program of instructions based on a script compiled from a source, txt file or lines.
/// Contains a list of instructions, ready to be executed one by one.
/// </summary>
public class Program
{
    public Program(Script script)
    {
        Script = script;
    }

    /// <summary>
    /// list of instruction to execute. Obtained after Parsing stage: Lexical analyse and Syntax analyse.
    /// </summary>
    public List<InstrBase> ListInstr { get; set; } = new List<InstrBase>();

    /// <summary>
    /// Script/source code text lines.
    /// From a text file or from a string stream.
    /// </summary>
    public Script Script { get; private set; }


    /// <summary>
    /// Find the last SetVar instr on the varname. on basic varname.
    /// a=12
    /// a=b
    /// 
    /// Not managed here, varname: A.Cell=12
    /// </summary>
    /// <param name="varname"></param>
    /// <returns></returns>
    public InstrSetVar FindLastVarSet(string varname)
    {
        // scan instr in reverse direction starting from the last one
        for (int i = ListInstr.Count - 1; i >= 0; i--)
        {
            InstrBase instr = ListInstr[i];
            if (instr.InstrType != InstrType.SetVar) continue;

            // ok, it's a setVar instr
            InstrSetVar instrSetVar = instr as InstrSetVar;

            // the left part should be a basic varname
            InstrObjectName instrObjectName= instrSetVar.InstrLeft as InstrObjectName;

            // not a basic varname
            if(instrObjectName== null) continue;

            // the current setVar left part does not match the var name
            if (!instrObjectName.ObjectName.Equals(varname, StringComparison.InvariantCultureIgnoreCase)) continue;

            // the right part is again a var name!
            if (instrSetVar.InstrRight.InstrType == InstrType.ObjectName)
                return FindLastVarSet((instrSetVar.InstrRight as InstrObjectName).ObjectName);
            
            // the current set var left part match the varname
            return instrSetVar;
        }

        // not found
        return null;
    }

}