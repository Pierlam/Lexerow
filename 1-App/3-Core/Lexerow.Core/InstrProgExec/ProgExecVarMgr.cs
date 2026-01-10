using Lexerow.Core.System;
using Lexerow.Core.System.GenDef;
using Lexerow.Core.System.InstrDef;
using Lexerow.Core.System.ProgExec;
using Lexerow.Core.System.ScriptDef;
using Lexerow.Core.Utils;

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
    /// For system var, name is simply a string.
    /// The value is a ValueBase object.
    /// </summary>
    public List<ProgExecSysVar> ListSysExecVar { get; private set; } = new List<ProgExecSysVar>();

    /// <summary>
    /// List of user defined execution variables.
    /// </summary>
    public List<ProgExecVar> ListExecVar { get; private set; } = new List<ProgExecVar>();


    /// <summary>
    /// Get the bool value of the var.
    /// </summary>
    /// <param name="result"></param>
    /// <param name="scriptToken"></param>
    /// <param name="varName"></param>
    /// <param name="value"></param>
    /// <returns></returns>
    public bool GetVarBoolValue(Result result, ScriptToken scriptToken, string varName, out bool value)
    {
        value = false;

        if (string.IsNullOrWhiteSpace(varName))
        {
            result.AddError(ErrorCode.ExecInstrVarNotFound, scriptToken);
            return false;
        }

        ProgExecVar progExecVar = FindVarByName(varName);
        if (progExecVar == null)
        {
            result.AddError(ErrorCode.ExecInstrVarNotFound, scriptToken);
            return false;
        }

        InstrValue instrValue = progExecVar.Value as InstrValue;
        if (instrValue==null)
        if (progExecVar.ObjectName.ReturnType != InstrReturnType.ValueBool)
        {
            result.AddError(ErrorCode.ExecInstrValueExpected, scriptToken);
            return false;
        }

        if(instrValue.ValueType != System.ValueType.Bool)
        {
            result.AddError(ErrorCode.ExecInstrBoolValueExpected, scriptToken);
            return false;
        }

        value = (instrValue.ValueBase as ValueBool).Val;
        return true;
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
    /// Only Possible to update an existing system var.
    /// For system var, name is simply a string.
    /// The value is a ValueBase object.
    /// </summary>
    /// <param name="instrVarLeft"></param>
    /// <param name="instrVarRight"></param>
    /// <returns></returns>
    public ProgExecSysVar UpdateSystemVar(string varName, ValueBase valueBase)
    {
        if (string.IsNullOrWhiteSpace(varName)) return null;
        if (valueBase == null) return null;

        // the var already defined ?
        ProgExecSysVar progExecSysVar = ListSysExecVar.FirstOrDefault(v => v.Name.Equals(varName,StringComparison.InvariantCultureIgnoreCase));

        if (progExecSysVar == null)
            return null;

        // check that the new value match the type
        if (!ValueUtils.TypeMatch(progExecSysVar.ValueBase, valueBase)) return null;


        // TODO: mettre dans un utils!  ValueUtils ?  
        if (!ValueUtils.CopyValue(progExecSysVar.ValueBase, valueBase))
            return null;

        return progExecSysVar;

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

        //--$StrCompareCaseSensitive=false
        progExecSysVar = new ProgExecSysVar(CoreInstr.SysVarStrCompareCaseSensitive, new ValueBool(false));
        ListSysExecVar.Add(progExecSysVar);
    }
}