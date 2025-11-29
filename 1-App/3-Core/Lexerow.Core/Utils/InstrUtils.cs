using Lexerow.Core.InstrProgExec;
using Lexerow.Core.System;
using Lexerow.Core.System.GenDef;
using Lexerow.Core.System.InstrDef;
using Lexerow.Core.System.ScriptDef;
using NPOI.SS.Formula.Functions;
using NPOI.XSSF.Streaming.Values;
using Org.BouncyCastle.Utilities.Collections;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static NPOI.HSSF.Util.HSSFColor;

namespace Lexerow.Core.Utils;

/// <summary>
/// Some useful functions around the InstrBase.
/// </summary>
public class InstrUtils
{
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
    public static bool GetStringFromInstr(Result result, bool isParser, Program program, InstrBase instr, out bool isValueOrVar, out string value)
    {
        isValueOrVar = false;
        value = string.Empty;

        //--is it an instr value?
        if (!GetValueStringFromInstrValue(result, isParser, instr, out isValueOrVar, out value))
            return false;
        if (isValueOrVar) return true;

        //--is it an instr var?
        if (!GetValueStringFromInstrVar(result, isParser, program, instr, out isValueOrVar, out value))
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
    public static bool GetIntFromInstr(Result result, bool isParser, Program program, InstrBase instr, out bool isValueOrVar, out int value)
    {
        isValueOrVar = false;
        value = 0;

        //--is it an instr value?
        if (!GetValueIntFromInstrValue(result, isParser, instr, out isValueOrVar, out value))
            return false;
        if (isValueOrVar) return true;

        //--is it an instr var?
        if (!GetValueIntFromInstrVar(result, isParser, program, instr, out isValueOrVar, out value))
            return false;
        
        return true;
    }

    public static bool GetValueIntFromInstrVar(Result result, bool isParser, Program program, InstrBase instr, out bool isVar, out int value)
    {
        isVar=false;
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
        if (!GetValueIntFromInstrValue(result, isParser, instrSetVar.InstrRight, out isVar, out value))
            return false;

        return true;
    }

    public static bool GetValueStringFromInstrVar(Result result, bool isParser, Program program, InstrBase instr, out bool isVar, out string value)
    {
        isVar = false;
        value = string.Empty;

        //--is it an var?
        InstrNameObject instrObjectName = instr as InstrNameObject;
        if (instrObjectName == null) return true;

        // check the final value of the var, can be a value, a fct call or a math expr
        InstrSetVar instrSetVar = program.FindLastVarSet(instrObjectName.Name);
        if (instrSetVar == null)
        {
            ErrorCode error = ErrorCode.ParserVarNotDefined;
            if(!isParser) error= ErrorCode.ExecInstrVarNotFound;
            result.AddError(error, instrObjectName.FirstScriptToken());
            return false;
        }
        isVar = true;

        //-the final var right instr is not a value?
        if (instrSetVar.InstrRight.InstrType != InstrType.Value)
            return true;

        //--is the instr right part a value?
        if (!GetValueStringFromInstrValue(result, isParser, instrSetVar.InstrRight, out isVar, out value))
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
    public static bool GetValueIntFromInstrValue(Result result, bool isParser, InstrBase instr, out bool isValue, out int value)
    {
        value = 0;
        isValue = false;

        var instrValue = instr as InstrValue;

        // not an instr value, nothing to do, bye
        if (instrValue == null) return true;

        // it's an InstrValue, type must be Int
        if (instrValue.ValueBase.ValueType != System.ValueType.Int)
        {
            ErrorCode error = ErrorCode.ParserValueIntExpected;
            if (!isParser) error = ErrorCode.ExecValueIntExpected;
            result.AddError(error, instrValue.FirstScriptToken().LineNum, instrValue.FirstScriptToken().ColNum, instrValue.FirstScriptToken().Value);
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
    public static bool GetValueStringFromInstrValue(Result result, bool isParser, InstrBase instr, out bool isValue, out string value)
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
            if(!isParser) error= ErrorCode.ExecValueStringExpected;
            result.AddError(error, instrValue.FirstScriptToken().LineNum, instrValue.FirstScriptToken().ColNum, instrValue.FirstScriptToken().Value);
            return false;
        }
        
        value = (instrValue.ValueBase as ValueString).Val;
        isValue= true;
        return true;
    }

    /// <summary>
    /// Merge the Minus char with the number, int or double.
    /// </summary>
    /// <param name="instrCharMinus"></param>
    /// <param name="instrValue"></param>
    /// <returns></returns>
    public static bool MergeInstrMinus(InstrCharMinus instrCharMinus, InstrValue instrValue)
    {
        if (instrCharMinus == null)return false;
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
        ValueInt valueInt= new ValueInt(initValue);
        ScriptToken scriptToken= new ScriptToken();
        scriptToken.Value= initValue.ToString();
        return new InstrValue(scriptToken, initValue);
    }

    public static bool IsValueInt(InstrBase instrBase)
    {
        if(instrBase == null)return false;
        if (instrBase.InstrType != InstrType.Value) return false;


        if ((instrBase as InstrValue).ValueBase.ValueType == System.ValueType.Int ) return true;
        return false;
    }

    public static bool IsValueDouble(InstrBase instrBase)
    {
        if (instrBase == null) return false;
        if (instrBase.InstrType != InstrType.Value) return false;


        if ((instrBase as InstrValue).ValueBase.ValueType == System.ValueType.Double) return true;
        return false;
    }
}
