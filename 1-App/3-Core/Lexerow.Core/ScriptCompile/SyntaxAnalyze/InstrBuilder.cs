using Lexerow.Core.System;
using Lexerow.Core.System.Compilator;
using Lexerow.Core.Utils;
using NPOI.Util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.ScriptCompile.SyntaxAnalyze;
public class InstrBuilder
{
    /// <summary>
    /// Create an instruction based on the script token.
    /// </summary>
    /// <param name="scriptToken"></param>
    /// <param name="instrBase"></param>
    /// <returns></returns>
    public static bool Build(ExecResult execResult, ScriptToken scriptToken, out InstrBase instrBase)
    {
        //--script token is a name/id
        if (scriptToken.ScriptTokenType == ScriptTokenType.Name)
        {
            // End
            if (scriptToken.Value.Equals("End", StringComparison.InvariantCultureIgnoreCase))
            {
                instrBase = new InstrEnd(scriptToken);
                return true;
            }

            // OpenExcel
            if (scriptToken.Value.Equals("OpenExcel", StringComparison.InvariantCultureIgnoreCase))
            {
                instrBase = new InstrOpenExcel(scriptToken);
                return true;
            }

            // OnExcel
            if (scriptToken.Value.Equals("OnExcel", StringComparison.InvariantCultureIgnoreCase))
            {
                instrBase = new InstrOnExcel(scriptToken);
                return true;
            }

            // OnSheet
            if (scriptToken.Value.Equals("OnSheet", StringComparison.InvariantCultureIgnoreCase))
            {
                instrBase = new InstrOnSheet(scriptToken);
                return true;
            }

            // ForEach
            if (scriptToken.Value.Equals("ForEach", StringComparison.InvariantCultureIgnoreCase))
            {
                instrBase = new InstrForEach(scriptToken);
                return true;
            }

            // Row
            if (scriptToken.Value.Equals("Row", StringComparison.InvariantCultureIgnoreCase))
            {
                instrBase = new InstrRow(scriptToken);
                return true;
            }

            // Next
            if (scriptToken.Value.Equals("Next", StringComparison.InvariantCultureIgnoreCase))
            {
                instrBase = new InstrNext(scriptToken);
                return true;
            }

            // Col
            if (scriptToken.Value.Equals("Col", StringComparison.InvariantCultureIgnoreCase))
            {
                instrBase = new InstrCol(scriptToken);
                return true;
            }

            // Cell
            if (scriptToken.Value.Equals("Cell", StringComparison.InvariantCultureIgnoreCase))
            {
                instrBase = new InstrCell(scriptToken);
                return true;
            }

            // if
            if (scriptToken.Value.Equals("if",StringComparison.InvariantCultureIgnoreCase))
            {
                instrBase = new InstrIf(scriptToken);
                return true;
            }

            // then
            if (scriptToken.Value.Equals("Then", StringComparison.InvariantCultureIgnoreCase))
            {
                instrBase = new InstrThen(scriptToken);
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
            if (scriptToken.Value.Equals(".", StringComparison.InvariantCultureIgnoreCase))
            {
                instrBase = new InstrDot(scriptToken);
                return true;
            }

            if (scriptToken.Value.Equals("+", StringComparison.InvariantCultureIgnoreCase))
            {
                instrBase = new InstrCharPlus(scriptToken);
                return true;
            }
            if (scriptToken.Value.Equals("-", StringComparison.InvariantCultureIgnoreCase))
            {
                instrBase = new InstrCharMinus(scriptToken);
                return true;
            }
            if (scriptToken.Value.Equals("*", StringComparison.InvariantCultureIgnoreCase))
            {
                instrBase = new InstrCharMul(scriptToken);
                return true;
            }

            if (scriptToken.Value.Equals("/", StringComparison.InvariantCultureIgnoreCase))
            {
                instrBase = new InstrCharDiv(scriptToken);
                return true;
            }

            if (scriptToken.Value.Equals("%", StringComparison.InvariantCultureIgnoreCase))
            {
                instrBase = new InstrCharPlus(scriptToken);
                return true;
            }

            execResult.AddError(new ExecResultError(ErrorCode.SyntaxAnalyzerTokenNotExpected, scriptToken.Value));
            instrBase = null;
            return false;
        }

        //--script token is a string
        if (scriptToken.ScriptTokenType == ScriptTokenType.String)
        {
            // remove double quote
            string val= StringUtils.RemoveStartEndDoubleQuote(scriptToken.Value);
            instrBase = new InstrConstValue(scriptToken, val);
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
        return false;
    }

    public static InstrSepComparison CreateSepComparison(ScriptToken scriptToken)
    {
        if (string.IsNullOrWhiteSpace(scriptToken.Value)) return null;
        return new InstrSepComparison(scriptToken);
    }

}
