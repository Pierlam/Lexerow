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
    public static bool GetValueIntFromInstrValue(InstrBase instr, int scriptLineNum, out ResultError error, out int value)
    {
        error = null;
        value = 0;

        var instrValue = instr as InstrValue;
        if (instrValue == null)
        {
            error = new ResultError(ErrorCode.ParserTokenExpected, scriptLineNum, instr.FirstScriptToken().ColNum, instr.FirstScriptToken().ToString());
            return false;
        }

        if (instrValue.ValueBase.ValueType != System.ValueType.Int)
        {
            error = new ResultError(ErrorCode.ParserValueIntExpected, instrValue.FirstScriptToken().LineNum, instrValue.FirstScriptToken().ColNum, instrValue.FirstScriptToken().Value);
            return false;
        }
        value = (instrValue.ValueBase as ValueInt).Val;
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
