using Lexerow.Core.InstrProgExec;
using Lexerow.Core.System;
using Lexerow.Core.System.InstrDef;
using Lexerow.Core.System.InstrDef.Object;
using Lexerow.Core.System.ScriptDef;
using OpenExcelSdk;

namespace Lexerow.Core.Utils;

/// <summary>
/// Some useful functions around the InstrBase.
/// </summary>
public class InstrUtils
{

    /// <summary>
    /// Is a var/fct call name?
    /// if it's the case, should exists in the list.
    /// </summary>
    /// <param name="result"></param>
    /// <param name="listVar"></param>
    /// <param name="instrBase"></param>
    /// <param name="instrNameObjectOut"></param>
    /// <returns></returns>
    public static bool CheckObjectName(Result result, List<InstrNameObject> listVar, InstrBase instrBase, out InstrNameObject instrNameObjectOut)
    {
        instrNameObjectOut = null;

        InstrNameObject instrNameObject = instrBase as InstrNameObject;
        if (instrNameObject == null) 
            // not the case
            return true;

        InstrNameObject instrObjectNameFound = listVar.FirstOrDefault(x => x.Name.Equals(instrNameObject.Name, StringComparison.InvariantCultureIgnoreCase));
        if (instrObjectNameFound== null)
        {
            // the var/fctcall should exists
            result.AddError(ErrorCode.ParserVarOrFctNameNotDefined, instrBase.FirstScriptToken());
            return false;
        }
        // push the existing one, because the return type is set
        instrNameObjectOut = instrObjectNameFound;
        return true;
    }


    /// <summary>
    /// Get the String value from the instruction, can be a Value or a Var.
    /// If it's a funct call, a math expr or a bool expr, can't return the value.
    /// </summary>
    /// <param name="result"></param>
    /// <param name="program"></param>
    /// <param name="instr"></param>
    /// <param name="isValueOrVar"></param>
    /// <param name="value"></param>
    /// <returns></returns>
    public static bool GetStringFromInstr(Result result, ProgExecVarMgr progExecVarMgr, InstrBase instr, out bool isValueOrVar, out string value)
    {
        isValueOrVar = false;
        value = string.Empty;

        //--is it an instr value?
        if (!GetStringFromInstrValue(result, false, instr, out isValueOrVar, out value))
            return false;
        if (isValueOrVar) return true;

        //--is it an instr var?
        if (!GetStringFromInstrVar(result, progExecVarMgr, instr, out isValueOrVar, out value))
            return false;

        return true;
    }

    public static bool GetFilenameOrExceFileFromInstr(Result result, ProgExecVarMgr progExecVarMgr, InstrBase instr, out string filename, out ExcelFile excelFile)
    {
        excelFile=null;
        bool isValueOrVar = false;

        //--is it an instr value?
        if (!GetStringFromInstrValue(result, false, instr, out isValueOrVar, out filename))
            return false;
        if (isValueOrVar) return true;

        //--is it a var?
        InstrNameObject instrObjectName = instr as InstrNameObject;
        if (instrObjectName == null) return true;

        // check the final value of the var, can be a value, a fct call or a math expr
        ProgExecVar progExecVar = progExecVarMgr.FindVarByName(instrObjectName.Name);
        if (progExecVar == null)
            return result.AddError(ErrorCode.ExecInstrVarNotFound, instrObjectName.FirstScriptToken());

        //--is the instr right part a value?
        if (!GetStringFromInstrValue(result, false, progExecVar.Value, out bool isValue, out filename))
            return false;

        if (isValue) return true;

        // is it an Excel file?
        InstrObjectExcelFile instrObjectExcelFile = progExecVar.Value as InstrObjectExcelFile;
        if(instrObjectExcelFile==null)
            // TODO: error?
            return true;

        filename = instrObjectExcelFile.Filename;
        excelFile = instrObjectExcelFile.ExcelFile;
        return true;
    }

    /// <summary>
    /// Get the first filename from the instruction ObjectSelectedFiles.
    /// </summary>
    /// <param name="progExecVarMgr"></param>
    /// <param name="instr"></param>
    /// <param name="value"></param>
    /// <returns></returns>
    public static bool GetFirstValueFromInstrObjectSelectedFiles(ProgExecVarMgr progExecVarMgr, InstrBase instr, out string value)
    {
        value = string.Empty;

        //--is it a var?
        InstrNameObject instrObjectName = instr as InstrNameObject;
        if (instrObjectName == null) return true;

        // check the final value of the var, can be a value, a fct call or a math expr
        ProgExecVar progExecVar = progExecVarMgr.FindVarByName(instrObjectName.Name);
        if (progExecVar == null)
            return false;

        //-the final var right instr is not a value?
        if (progExecVar.Value.InstrType != InstrType.ObjectSelectedFiles)
            return true;

        InstrObjectSelectedFiles instrObjectSelectedFiles = progExecVar.Value as InstrObjectSelectedFiles;
        if (instrObjectSelectedFiles.ListObjectSelectedFile.Count == 0)
            return false;

        value= instrObjectSelectedFiles.ListObjectSelectedFile[0].Filename;
        return true;
    }

    /// <summary>
    /// Get the String value from the instruction, can be a Value or a Var.
    /// If it's a funct call, a math expr or a bool expr, can't return the value.
    /// </summary>
    /// <param name="result"></param>
    /// <param name="program"></param>
    /// <param name="instr"></param>
    /// <param name="isValueOrVar"></param>
    /// <param name="value"></param>
    /// <returns></returns>
    public static bool GetStringFromInstrParser(Result result, Program program, InstrBase instr, out bool isValueOrVar, out string value)
    {
        isValueOrVar = false;
        value = string.Empty;

        //--is it an instr value?
        if (!GetStringFromInstrValue(result, true, instr, out isValueOrVar, out value))
            return false;
        if (isValueOrVar) return true;

        //--is it an instr var?
        if (!GetStringFromInstrVarParser(result, program, instr, out isValueOrVar, out value))
            return false;

        return true;
    }


    /// <summary>
    /// Get the int value from the instruction, can be a Value or a Var.
    /// If it's a funct call, a math expr or a bool expr, can't return the value.
    /// </summary>
    /// <param name="result"></param>
    /// <param name="program"></param>
    /// <param name="instr"></param>
    /// <param name="isValueOrVar"></param>
    /// <param name="value"></param>
    /// <returns></returns>
    public static bool GetIntFromInstrExec(Result result, ProgExecVarMgr progExecVarMgr, InstrBase instr, out bool isValueOrVar, out int value)
    {
        isValueOrVar = false;
        value = 0;

        //--is it an instr value?
        if (!GetIntFromInstrValueParser(result, instr, out isValueOrVar, out value))
            return false;
        if (isValueOrVar) return true;

        //--is it an instr var?
        if (!GetIntFromInstrVarExec(result, progExecVarMgr, instr, out isValueOrVar, out value))
            return false;

        return true;
    }

    /// <summary>
    /// Get the int value from the instruction, can be a Value or a Var.
    /// If it's a funct call, a math expr or a bool expr, can't return the value.
    /// </summary>
    /// <param name="result"></param>
    /// <param name="program"></param>
    /// <param name="instr"></param>
    /// <param name="isValueOrVar"></param>
    /// <param name="value"></param>
    /// <returns></returns>
    public static bool GetIntFromInstrParser(Result result, Program program, InstrBase instr, out bool isValueOrVar, out int value)
    {
        isValueOrVar = false;
        value = 0;

        //--is it an instr value?
        if (!GetIntFromInstrValueParser(result, instr, out isValueOrVar, out value))
            return false;
        if (isValueOrVar) return true;

        //--is it an instr var?
        if (!GetIntFromInstrVar(result, true, program, instr, out isValueOrVar, out value))
            return false;

        return true;
    }

    public static bool GetIntFromInstrVar(Result result, bool isParser, Program program, InstrBase instr, out bool isVar, out int value)
    {
        isVar = false;
        value = 0;

        //--is it an var?
        InstrNameObject instrObjectName = instr as InstrNameObject;
        if (instrObjectName == null) return true;

        // check the final value of the var, can be a value, a fct call or a math expr
        InstrSetVar instrSetVar = program.FindLastVarSet(instrObjectName.Name);
        if (instrSetVar == null)
        {
            ErrorCode error = ErrorCode.ParserVarNotDefined;
            if (!isParser) error = ErrorCode.ExecInstrVarNotFound;
            result.AddError(error, instrObjectName.FirstScriptToken());
            return false;
        }
        isVar = true;

        //-the final var right instr is not a value?
        if (instrSetVar.InstrRight.InstrType != InstrType.Value)
            return true;

        //--is the instr right part a value?
        if (!GetIntFromInstrValueParser(result, instrSetVar.InstrRight, out isVar, out value))
            return false;

        return true;
    }

    public static bool GetIntFromInstrVarExec(Result result, ProgExecVarMgr progExecVarMgr, InstrBase instr, out bool isVar, out int value)
    {
        isVar = false;
        value = 0;

        //--is it an var?
        InstrNameObject instrObjectName = instr as InstrNameObject;
        if (instrObjectName == null) return true;

        // check the final value of the var, can be a value, a fct call or a math expr
        ProgExecVar progExecVar = progExecVarMgr.FindVarByName(instrObjectName.Name);
        if (progExecVar == null)
            return result.AddError(ErrorCode.ExecInstrVarNotFound, instrObjectName.FirstScriptToken());

        isVar = true;

        //-the final var right instr is not a value?
        if (progExecVar.Value.InstrType != InstrType.Value)
            return true;

        //--is the instr right part a value?
        if (!GetIntFromInstrValueParser(result, progExecVar.Value, out isVar, out value))
            return false;

        return true;
    }

    /// <summary>
    /// Get the value of the var instruction, if it's a var, else return null.
    /// </summary>
    /// <param name="progExecVarMgr"></param>
    /// <param name="instr"></param>
    /// <returns></returns>
    public static InstrBase GetValueFromInstrVarParser(Program program, InstrBase instr)
    {

        //--is it a var?
        InstrNameObject instrObjectName = instr as InstrNameObject;
        if (instrObjectName == null) return null;

        InstrSetVar instrSetVar = program.FindLastVarSet(instrObjectName.Name);
        if (instrSetVar == null) return null;

        return instrSetVar.InstrRight;
    }

    public static bool GetStringFromInstrVarParser(Result result, Program program, InstrBase instr, out bool isVar, out string value)
    {
        isVar = false;
        value = string.Empty;

        //--is it a var?
        InstrNameObject instrObjectName = instr as InstrNameObject;
        if (instrObjectName == null) return true;

        // check the final value of the var, can be a value, a fct call or a math expr
        InstrSetVar instrSetVar = program.FindLastVarSet(instrObjectName.Name);
        if (instrSetVar == null)
            return result.AddError(ErrorCode.ParserVarNotDefined, instrObjectName.FirstScriptToken());

        isVar = true;

        //-the final var right instr is not a value?
        if (instrSetVar.InstrRight.InstrType != InstrType.Value)
            return true;

        //--is the instr right part a value?
        if (!GetStringFromInstrValue(result, true, instrSetVar.InstrRight, out isVar, out value))
            return false;

        return true;
    }

    /// <summary>
    /// Get the object excel file from the instruction, if it's a var, else return null.
    /// </summary>
    /// <param name="result"></param>
    /// <param name="progExecVarMgr"></param>
    /// <param name="instr"></param>
    /// <param name="value"></param>
    /// <returns></returns>
    public static bool GetObjectExcelFileFromInstrVar(Result result, ProgExecVarMgr progExecVarMgr, InstrBase instr,  out InstrObjectExcelFile value)
    {
        value = null;

        //--is it a var?
        InstrNameObject instrObjectName = instr as InstrNameObject;
        if (instrObjectName == null) return true;

        // check the final value of the var, can be a value, a fct call or a math expr
        ProgExecVar progExecVar = progExecVarMgr.FindVarByName(instrObjectName.Name);
        if (progExecVar == null)
            return result.AddError(ErrorCode.ExecInstrVarNotFound, instrObjectName.FirstScriptToken());

        // is it an Excel file?
        InstrObjectExcelFile instrObjectExcelFile = progExecVar.Value as InstrObjectExcelFile;
        if (instrObjectExcelFile == null)
            // TODO: error?
            return true;

        value= instrObjectExcelFile;
        return true;
    }

    public static bool GetStringFromInstrVar(Result result, ProgExecVarMgr progExecVarMgr, InstrBase instr, out bool isVar, out string value)
    {
        isVar = false;
        value = string.Empty;

        //--is it a var?
        InstrNameObject instrObjectName = instr as InstrNameObject;
        if (instrObjectName == null) return true;

        // check the final value of the var, can be a value, a fct call or a math expr
        ProgExecVar progExecVar = progExecVarMgr.FindVarByName(instrObjectName.Name);
        if (progExecVar == null)
            return result.AddError(ErrorCode.ExecInstrVarNotFound, instrObjectName.FirstScriptToken());

        isVar = true;

        //-the final var right instr is not a value?
        if (progExecVar.Value.InstrType != InstrType.Value)
            return true;

        //--is the instr right part a value?
        if (!GetStringFromInstrValue(result, false, progExecVar.Value, out isVar, out value))
            return false;

        return true;
    }

    /// <summary>
    /// If it's not an instr value, not checked, ok.
    /// </summary>
    /// <param name="result"></param>
    /// <param name="instr"></param>
    /// <param name="isValue"></param>
    /// <param name="value"></param>
    /// <returns></returns>
    public static bool GetIntFromInstrValueParser(Result result, InstrBase instr, out bool isValue, out int value)
    {
        value = 0;
        isValue = false;

        var instrValue = instr as InstrValue;

        // not an instr value, nothing to do, bye
        if (instrValue == null) return true;

        // it's an InstrValue, type must be Int
        if (instrValue.ValueBase.ValueType != System.ValueType.Int)
        {
            result.AddError(ErrorCode.ParserValueIntExpected, instrValue.FirstScriptToken().LineNum, instrValue.FirstScriptToken().ColNum, instrValue.FirstScriptToken().Value);
            return false;
        }
        value = (instrValue.ValueBase as ValueInt).Val;
        isValue = true;
        return true;
    }

    /// <summary>
    /// If it's not an instr value, not checked, ok.
    /// </summary>
    /// <param name="result"></param>
    /// <param name="instr"></param>
    /// <param name="isValue"></param>
    /// <param name="value"></param>
    /// <returns></returns>
    public static bool GetStringFromInstrValue(Result result, bool isParser, InstrBase instr, out bool isValue, out string value)
    {
        value = string.Empty;
        isValue = false;

        var instrValue = instr as InstrValue;

        // not an instr value, nothing to do, bye
        if (instrValue == null) return true;

        // it's an InstrValue, type must be Int
        if (instrValue.ValueBase.ValueType != System.ValueType.String)
        {
            ErrorCode error = ErrorCode.ParserValueStringExpected;
            if (!isParser) error = ErrorCode.ExecValueStringExpected;
            result.AddError(error, instrValue.FirstScriptToken().LineNum, instrValue.FirstScriptToken().ColNum, instrValue.FirstScriptToken().Value);
            return false;
        }

        value = (instrValue.ValueBase as ValueString).Val;
        isValue = true;
        return true;
    }

    public static void ClearIsExecuted(InstrBase instr)
    {
        if (instr == null) return;
        instr.IsExecuted = false;

        InstrBoolExpr instrBoolExpr = instr as InstrBoolExpr;
        if(instrBoolExpr!=null)
        {
            foreach(InstrBase instrInner in instrBoolExpr.ListOperand)
            {
                ClearIsExecuted(instrInner);
            }
            return;
        }

        InstrComparison instrComparison = instr as InstrComparison;
        if (instrComparison != null) 
        {
            ClearIsExecuted(instrComparison.OperandLeft);
            ClearIsExecuted(instrComparison.OperandRight);
            return;
        }


        //TODO: InstrMathExpr 
    }

    /// <summary>
    /// Return the type of the instr value.
    /// </summary>
    /// <param name="instrValue"></param>
    /// <returns></returns>
    public static InstrReturnType GetReturnType(InstrValue instrValue)
    {
        if(instrValue==null)return InstrReturnType.Nothing;

        if (instrValue.ValueType == System.ValueType.String) return InstrReturnType.ValueString;
        if (instrValue.ValueType == System.ValueType.Int) return InstrReturnType.ValueInt;
        if (instrValue.ValueType == System.ValueType.Double) return InstrReturnType.ValueDouble;
        if (instrValue.ValueType == System.ValueType.Bool) return InstrReturnType.ValueBool;

        if (instrValue.ValueType == System.ValueType.DateOnly) return InstrReturnType.ValueDateOnly;
        if (instrValue.ValueType == System.ValueType.DateTime) return InstrReturnType.ValueDateTime;
        if (instrValue.ValueType == System.ValueType.TimeOnly) return InstrReturnType.ValueTimeOnly;

        return InstrReturnType.Nothing;
    }


    /// <summary>
    /// Create a and or Or instruction.
    /// </summary>
    /// <param name="scriptToken"></param>
    /// <returns></returns>
    public static InstrBase CreateBoolOperator(ScriptToken scriptToken)
    {
        if (scriptToken.Value.Equals("And", StringComparison.InvariantCultureIgnoreCase))
            return new InstrAnd(scriptToken);

        if (scriptToken.Value.Equals("Or", StringComparison.InvariantCultureIgnoreCase))
            return new InstrOr(scriptToken);

        return null;
    }

    /// <summary>
    /// Is the instruction need to be executed?
    /// It's the case for: function call, comparison, bool and math expression.
    /// </summary>
    /// <param name="instrBase"></param>
    /// <returns></returns>
    public static bool NeedToBeExecuted(InstrBase instrBase)
    {
        if (instrBase == null) return false;

        if (instrBase.IsFunctionCall) return true;
        if (instrBase.InstrType == InstrType.Comparison) return true;
        if (instrBase.InstrType == InstrType.BoolExpr) return true;
        if (instrBase.InstrType == InstrType.MathExpr) return true;

        // no
        return false;
    }

    /// <summary>
    /// Merge the Minus char with the number, int or double.
    /// </summary>
    /// <param name="instrCharMinus"></param>
    /// <param name="instrValue"></param>
    /// <returns></returns>
    public static bool MergeInstrMinus(InstrCharMinus instrCharMinus, InstrValue instrValue)
    {
        if (instrCharMinus == null) return false;
        if (instrValue == null) return false;

        if (instrValue.ValueBase.ValueType == System.ValueType.Int)
        {
            ValueInt valueInt = instrValue.ValueBase as ValueInt;
            valueInt.Val = -1 * valueInt.Val;
            instrValue.RawValue = valueInt.Val.ToString();
            return true;
        }

        ValueDouble valueDouble = instrValue.ValueBase as ValueDouble;
        valueDouble.Val = -1 * valueDouble.Val;
        instrValue.RawValue = valueDouble.Val.ToString();
        return true;
    }
    public static InstrValue CreateInstrValueInt(int initValue)
    {
        ValueInt valueInt = new ValueInt(initValue);
        ScriptToken scriptToken = new ScriptToken();
        scriptToken.Value = initValue.ToString();
        return new InstrValue(scriptToken, initValue);
    }

    public static bool IsValueInt(InstrBase instrBase)
    {
        if (instrBase == null) return false;
        if (instrBase.InstrType != InstrType.Value) return false;

        if ((instrBase as InstrValue).ValueBase.ValueType == System.ValueType.Int) return true;
        return false;
    }

    public static bool IsValueDouble(InstrBase instrBase)
    {
        if (instrBase == null) return false;
        if (instrBase.InstrType != InstrType.Value) return false;

        if ((instrBase as InstrValue).ValueBase.ValueType == System.ValueType.Double) return true;
        return false;
    }

    public static bool IsValueString(InstrBase instrBase)
    {
        if (instrBase == null) return false;
        if (instrBase.InstrType != InstrType.Value) return false;

        if ((instrBase as InstrValue).ValueBase.ValueType == System.ValueType.String) return true;
        return false;
    }

}