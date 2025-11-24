using Lexerow.Core.System;
using Lexerow.Core.System.GenDef;
using Lexerow.Core.System.InstrDef;
using Lexerow.Core.System.ScriptDef;
using NPOI.SS.Formula.Functions;
using Org.BouncyCastle.Utilities.Collections;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Utils;

/// <summary>
/// Some useful functions around the InstrBase.
/// </summary>
public class InstrUtils
{
    /// <summary>
    /// Check the instr, should be string value or a var or a fct call to a string value.
    /// </summary>
    /// <param name="result"></param>
    /// <param name="listVar"></param>
    /// <param name="program"></param>
    /// <param name="instr"></param>
    /// <param name="instrToCheck"></param>
    /// <returns></returns>
    public static bool ChekInstrString(Result result, List<InstrObjectName> listVar, Program program, InstrBase instrToCheck, out bool valueSet, out string value)
    {
        valueSet = false;
        value = string.Empty;

        //--is it an instr value?

        InstrValue instrValue = instrToCheck as InstrValue;
        if (instrValue != null)
        {
            if (!GetValueStringFromInstrValue(result, instrToCheck, instrToCheck.FirstScriptToken().LineNum, out value))
                return false;
            valueSet = true;
            return true;
        }

        //--is it an var?
        InstrObjectName instrObjectName = instrToCheck as InstrObjectName;
        if (instrObjectName != null)
        {
            // check the final value of the var, can be a value, a fct call or a math expr
            InstrSetVar instrSetVar = program.FindLastVarSet(instrObjectName.ObjectName);
            if (instrSetVar == null)
            {
                result.AddError(ErrorCode.ParserVarNotDefined, instrObjectName.FirstScriptToken());
                return false;
            }

            //-the final var right instr is a value?
            if (instrSetVar.InstrRight.InstrType == InstrType.Value)
            {
                instrValue = instrSetVar.InstrRight as InstrValue;
                if (!GetValueStringFromInstrValue(result, instrValue, instrValue.FirstScriptToken().LineNum, out value))
                    return false;

                valueSet = true;
                return true;
            }

            //-the final var right instr is a fct call? a=fct()
            // TODO: to implement

            //-the final var right instr is a math expr? a=12+5
            // TODO: to implement

            result.AddError(ErrorCode.ParserVarWrongRightPart, instrToCheck.ListScriptToken[0]);
            return false;
        }

        //--is it a fct call? exp: a=fct()
        // TODO: to implement

        //--is it a math expr? exp: a=12+5
        // TODO: to implement

        result.AddError(ErrorCode.ParserValueIntExpected, instrToCheck.ListScriptToken[0]);
        return false;
    }

    /// <summary>
    /// Check that the instr is an int value
    /// </summary>
    /// <param name="result"></param>
    /// <param name="program"></param>
    /// <param name="instrToCheck"></param>
    /// <param name="valueSet"></param>
    /// <param name="value"></param>
    /// <returns></returns>
    public static bool CheckInstrInt(Result result, Program program, InstrBase instrToCheck, out bool valueSet, out int value)
    {
        valueSet = false;
        value = 0;

        //--is it an instr value?
        InstrValue instrValue = instrToCheck as InstrValue;
        if (instrValue != null)
        {
            if (!GetValueIntFromInstrValue(result, instrToCheck, instrToCheck.FirstScriptToken().LineNum, out value))
                return false;
            valueSet = true;
            return true;
        }

        //--is it an var?
        InstrObjectName instrObjectName = instrToCheck as InstrObjectName;
        if (instrObjectName != null)
        {
            // check the final value of the var, can be a value, a fct call or a math expr
            InstrSetVar instrSetVar = program.FindLastVarSet(instrObjectName.ObjectName);
            if (instrSetVar == null)
            {
                result.AddError(ErrorCode.ParserVarNotDefined, instrObjectName.FirstScriptToken());
                return false;
            }

            //-the final var right instr is a value?
            if (instrSetVar.InstrRight.InstrType == InstrType.Value)
            {
                instrValue = instrSetVar.InstrRight as InstrValue;
                if (!GetValueIntFromInstrValue(result, instrValue, instrValue.FirstScriptToken().LineNum, out value))
                    return false;

                valueSet = true;
                return true;
            }

            //-the final var right instr is a fct call? a=fct()
            // TODO: to implement

            //-the final var right instr is a math expr? a=12+5
            // TODO: to implement

            result.AddError(ErrorCode.ParserVarWrongRightPart, instrToCheck.ListScriptToken[0]);
            return false;
        }

        //--is it a fct call? exp: a=fct()
        // TODO: to implement

        //--is it a math expr? exp: a=12+5
        // TODO: to implement

        result.AddError(ErrorCode.ParserValueIntExpected, instrToCheck.ListScriptToken[0]);
        return false;
    }

    /// <summary>
    /// Return the value int, if it's a case.
    /// </summary>
    /// <param name="instrBase"></param>
    /// <param name="scriptLineNum"></param>
    /// <param name="error"></param>
    /// <param name="value"></param>
    /// <returns></returns>
    public static bool GetValueInt(InstrBase instrBase, int scriptLineNum, out int value)
    {
        value = 0;

        if(instrBase==null)
            return false;

        InstrValue instrValue = instrBase as InstrValue;
        if (instrValue == null) 
            return false;

        if (instrValue.ValueBase.ValueType != System.ValueType.Int)
            return false;

        value = (instrValue.ValueBase as ValueInt).Val;
        return true;
    }

    /// <summary>
    /// Return the value int from the InstrValue parameter.
    /// </summary>
    /// <param name="instrBase"></param>
    /// <param name="scriptLineNum"></param>
    /// <param name="error"></param>
    /// <param name="value"></param>
    /// <returns></returns>
    public static bool GetValueIntFromInstrValue(Result result, InstrBase instr, int scriptLineNum, out int value)
    {
        value = 0;

        var instrValue = instr as InstrValue;
        if (instrValue == null)
        {
            result.AddError(ErrorCode.ParserTokenExpected, scriptLineNum, instr.FirstScriptToken().ColNum, instr.FirstScriptToken().ToString());
            return false;
        }

        if (instrValue.ValueBase.ValueType != System.ValueType.Int)
        {
            result.AddError(ErrorCode.ParserValueIntExpected, instrValue.FirstScriptToken().LineNum, instrValue.FirstScriptToken().ColNum, instrValue.FirstScriptToken().Value);
            return false;
        }
        value = (instrValue.ValueBase as ValueInt).Val;
        return true;
    }

    /// <summary>
    /// Return the value string from the InstrValue parameter.
    /// </summary>
    /// <param name="instrBase"></param>
    /// <param name="scriptLineNum"></param>
    /// <param name="error"></param>
    /// <param name="value"></param>
    /// <returns></returns>
    public static bool GetValueStringFromInstrValue(Result result, InstrBase instr, int scriptLineNum, out string value)
    {
        value = string.Empty;

        var instrValue = instr as InstrValue;
        if (instrValue == null)
        {
            result.AddError(ErrorCode.ParserTokenExpected, scriptLineNum, instr.FirstScriptToken().ColNum, instr.FirstScriptToken().ToString());
            return false;
        }

        if (instrValue.ValueBase.ValueType != System.ValueType.String)
        {
            result.AddError(ErrorCode.ParserValueStringExpected, instrValue.FirstScriptToken().LineNum, instrValue.FirstScriptToken().ColNum, instrValue.FirstScriptToken().Value);
            return false;
        }
        value = (instrValue.ValueBase as ValueString).Val;
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
