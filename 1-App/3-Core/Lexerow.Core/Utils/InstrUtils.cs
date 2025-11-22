using Lexerow.Core.System;
using Lexerow.Core.System.GenDef;
using Lexerow.Core.System.ScriptDef;
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

        InstrValue instrConstValue = instrBase as InstrValue;
        if (instrConstValue == null) 
            return false;

        if (instrConstValue.ValueBase.ValueType != System.ValueType.Int)
            return false;

        value = (instrConstValue.ValueBase as ValueInt).Val;
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
    public static bool GetValueIntFromInstrValue(InstrBase instr, int scriptLineNum, out ExecResultError error, out int value)
    {
        error = null;
        value = 0;

        var instrConstValue = instr as InstrValue;
        if (instrConstValue == null)
        {
            error = new ExecResultError(ErrorCode.ParserTokenExpected, scriptLineNum, instr.FirstScriptToken().ColNum, instr.FirstScriptToken().ToString());
            return false;
        }

        if (instrConstValue.ValueBase.ValueType != System.ValueType.Int)
        {
            error = new ExecResultError(ErrorCode.ParserConstIntValueExpected, instrConstValue.FirstScriptToken().LineNum, instrConstValue.FirstScriptToken().ColNum, instrConstValue.FirstScriptToken().Value);
            return false;
        }
        value = (instrConstValue.ValueBase as ValueInt).Val;
        return true;
    }

    public static InstrValue CreateInstrValueInt(int initValue)
    {
        ValueInt valueInt= new ValueInt(initValue);
        ScriptToken scriptToken= new ScriptToken();
        scriptToken.Value= initValue.ToString();
        return new InstrValue(scriptToken, initValue);
    }

}
