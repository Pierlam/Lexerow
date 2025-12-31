using Lexerow.Core.System;
using Lexerow.Core.System.GenDef;
using Lexerow.Core.System.InstrDef;
using Lexerow.Core.System.ProgExec;

namespace Lexerow.Core.InstrProgExec;

/// <summary>
/// Program execution variables manager.
/// </summary>
public class ProgExecVarMgr
{
    public ProgExecVarMgr()
    {
        BuildSysVar();
    }

    /// <summary>
    /// List of system/built-in variables.
    /// </summary>
    public List<ProgExecSysVar> ListSysExecVar { get; private set; } = new List<ProgExecSysVar>();

    /// <summary>
    /// List of user defined execution variables.
    /// </summary>
    public List<ProgExecVar> ListExecVar { get; private set; } = new List<ProgExecVar>();

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

    public string GetProgExecSysVarAsString(string name)
    {
        ProgExecSysVar progExecSysVar = GetProgExecSysVar(name);
        if (progExecSysVar == null) return string.Empty;

        if (progExecSysVar.ValueBase.ValueType != System.ValueType.String) return string.Empty;
        return (progExecSysVar.ValueBase as ValueString).Val;
    }

    public int GetProgExecSysVarAsInt(string name)
    {
        ProgExecSysVar progExecSysVar = GetProgExecSysVar(name);
        if (progExecSysVar == null) return 0;

        if (progExecSysVar.ValueBase.ValueType != System.ValueType.Int) return 0;
        return (progExecSysVar.ValueBase as ValueInt).Val;
    }

    public bool GetProgExecSysVarAsBool(string name)
    {
        ProgExecSysVar progExecSysVar = GetProgExecSysVar(name);
        if (progExecSysVar == null) return false;

        if (progExecSysVar.ValueBase.ValueType != System.ValueType.Bool) return false;
        return (progExecSysVar.ValueBase as ValueBool).Val;
    }

    public ProgExecSysVar GetProgExecSysVar(string name)
    {
        if (string.IsNullOrWhiteSpace(name)) return null;
        return ListSysExecVar.FirstOrDefault(x => x.Name.Equals(name, StringComparison.InvariantCultureIgnoreCase));
    }

    private void BuildSysVar()
    {
        ProgExecSysVar progExecSysVar;

        //--$DateFormat
        progExecSysVar = new ProgExecSysVar(CoreInstr.SysVarDateFormatName, new ValueString("d/m/yyyy"));
        ListSysExecVar.Add(progExecSysVar);

        //--$DateTimeFormat
        progExecSysVar = new ProgExecSysVar(CoreInstr.SysVarDateTimeFormatName, new ValueString("dd/mm/yyyy\\ hh:mm"));
        ListSysExecVar.Add(progExecSysVar);

        //--$TimeFormat
        progExecSysVar = new ProgExecSysVar(CoreInstr.SysVarTimeFormatName, new ValueString("hh:mm:ss"));
        ListSysExecVar.Add(progExecSysVar);

        //--$CurrencyFormat
        progExecSysVar = new ProgExecSysVar(CoreInstr.SysVarCurrencyFormatName, new ValueString("#,##0.00\\ \"€\""));
        ListSysExecVar.Add(progExecSysVar);

        //--$ForceDateFormat=false
        progExecSysVar = new ProgExecSysVar(CoreInstr.SysVarForceDateFormatName, new ValueBool(false));
        ListSysExecVar.Add(progExecSysVar);
    }
}