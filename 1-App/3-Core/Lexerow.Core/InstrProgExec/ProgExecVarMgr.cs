using Lexerow.Core.System;
using Lexerow.Core.System.InstrDef;

namespace Lexerow.Core.InstrProgExec;

/// <summary>
/// Program execution variables manager.
/// </summary>
public class ProgExecVarMgr
{
    /// <summary>
    /// List of defined execution variables.
    /// </summary>
    public List<ProgExecVar> ListExecVar { get; private set; } = new List<ProgExecVar>();

    public void Add(ProgExecVar progExecVar)
    {
        ListExecVar.Add(progExecVar);
    }

    /// <summary>
    /// Find the last var. Usefull if the value of the var is a var.
    /// exp a=b
    /// b=12
    /// > will return the var b.
    /// </summary>
    /// <param name="varname"></param>
    /// <returns></returns>
    public ProgExecVar FindLastInnerVarByName(string varname)
    {
        ProgExecVar currProgExecVar = null;

        while (true)
        {
            currProgExecVar = ListExecVar.FirstOrDefault(v => v.NameEquals(varname));
            if (currProgExecVar == null) return null;

            var v = currProgExecVar.Value as InstrNameObject;
            if (v == null) return currProgExecVar;

            // now find the var value
            varname = v.Name;
        }
    }
}