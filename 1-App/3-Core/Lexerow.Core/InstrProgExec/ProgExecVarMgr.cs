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
    /// Create a new var, or update if the var already exists (left part).
    /// instrVarLeft= instrVarRight
    /// a=12
    /// A.Cell=12
    /// </summary>
    /// <param name="instrVarLeft"></param>
    /// <param name="instrVarRight"></param>
    /// <returns></returns>
    public ProgExecVar CreateOrUpdateVar(InstrBase instrVarLeft, InstrBase instrVarRight)
    {
        // the var already defined ?
        ProgExecVar progExecVar = ListExecVar.FirstOrDefault(v => v.AreEqual(instrVarLeft));

        if (progExecVar != null) 
        {
            progExecVar.Value = instrVarRight;
            return progExecVar;
        }

        // create the var
        progExecVar = new ProgExecVar(instrVarLeft, instrVarRight);
        ListExecVar.Add(progExecVar);
        return progExecVar;
    }

    /// <summary>
    /// Return the var object by the var name.
    /// Only if the var name is basic, exp: a=12
    /// return null in others cases, exp A.Cell=12
    /// </summary>
    /// <param name="varname"></param>
    /// <returns></returns>
    public ProgExecVar FindVarByName(string varname)
    {
        if (string.IsNullOrWhiteSpace(varname)) return null;

        return ListExecVar.FirstOrDefault(v => v.NameEquals(varname));
    }

    /// <summary>
    /// Find the last var. Useful if the value of the var is a var.
    /// exp a=b
    /// b=12
    /// > will return the var b.
    /// TODO:NOT NORMAL! REMOVE IT
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