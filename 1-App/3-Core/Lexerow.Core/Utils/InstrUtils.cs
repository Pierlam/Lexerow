using Lexerow.Core.System;
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
    public static bool GetConstValueInt(InstrBase instrBase, int scriptLineNum, out ExecResultError error, out int value)
    {
        error = null;
        value = 0;

        if(instrBase==null)
        {
            error = new ExecResultError(ErrorCode.ParserTokenExpected, scriptLineNum, 0, string.Empty);
            return false;
        }

        InstrConstValue instrConstValue = instrBase as InstrConstValue;
        if (instrConstValue == null) 
        {
            error = new ExecResultError(ErrorCode.ParserConstStringValueExpected, instrBase.FirstScriptToken().LineNum, instrBase.FirstScriptToken().ColNum, instrBase.FirstScriptToken().Value);
            return false;
        }

        if (instrConstValue.ValueBase.ValueType != System.ValueType.Int)
        {
            error = new ExecResultError(ErrorCode.ParserConstIntValueExpected, instrBase.FirstScriptToken().LineNum, instrBase.FirstScriptToken().ColNum, instrBase.FirstScriptToken().Value);
            return false;
        }
        value = (instrConstValue.ValueBase as ValueInt).Val;
        return true;
    }
}
