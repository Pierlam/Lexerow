using Lexerow.Core.System;
using Lexerow.Core.System.Compilator;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Scripts.SyntaxAnalyze;
public class InstrBuilder
{
    /// <summary>
    /// Create an instruction based on the script token.
    /// </summary>
    /// <param name="scriptToken"></param>
    /// <param name="instrBase"></param>
    /// <returns></returns>
    public static bool Do(ExecResult execResult, ScriptToken scriptToken, out InstrBase instrBase)
    {
        //--script token is a name/id
        if (scriptToken.ScriptTokenType == ScriptTokenType.Name)
        {
            // OpenExcel
            if (scriptToken.Value.Equals("OpenExcel", StringComparison.InvariantCultureIgnoreCase))
            {
                instrBase = new InstrOpenExcel(scriptToken);
                return true;
            }

            // OnExcel
            // TODO:

            // OnSheet
            // TODO:

            // ForEach
            // TODO:

            // Row
            // TODO:

            // Col
            // TODO:

            // if
            if (scriptToken.Value.Equals("if",StringComparison.InvariantCultureIgnoreCase))
            {
                instrBase = null;
                return true;
            }

            // then
            if (scriptToken.Value.Equals("Then", StringComparison.InvariantCultureIgnoreCase))
            {
                instrBase = null;
                return true;
            }

            // ExcelCol, exp: A
            // TODO:


            // ExcelCellAddress, exp: A1
            // TODO:

            // if it is not a known keyword, it's an object name, can be: a variable or a user defined function 
            instrBase = new InstrObjectName(scriptToken);
            return true;
        }

        //--script token is a separator
        if (scriptToken.ScriptTokenType == ScriptTokenType.Separator)
        {
            if (scriptToken.Value.Equals("(", StringComparison.InvariantCultureIgnoreCase))
            {
                instrBase = new InstrOpenBracket(scriptToken);
                return true;
            }
            if (scriptToken.Value.Equals(")", StringComparison.InvariantCultureIgnoreCase))
            {
                // TODO: needed? 
                //instrBase = new InstrCloseBrace(scriptToken);
                instrBase = null;
                return true;
            }


        }

        //--script token is a string
        if (scriptToken.ScriptTokenType == ScriptTokenType.String)
        {
            instrBase = new InstrConstValue(scriptToken, scriptToken.Value);
            return true;
        }

        //--script token is a int
        if (scriptToken.ScriptTokenType == ScriptTokenType.Integer)
        {
            instrBase = new InstrConstValue(scriptToken, scriptToken.ValueInt);
            return true;
        }

        //--script token is a double
        if (scriptToken.ScriptTokenType == ScriptTokenType.Double)
        {
            instrBase = new InstrConstValue(scriptToken, scriptToken.ValueDouble);
            return true;
        }


        execResult.AddError(new ExecResultError(ErrorCode.SyntaxAnalyzerTokenNotExpected, scriptToken.Value));
        instrBase = null;
        return true;
    }
}
