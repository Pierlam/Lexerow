using Lexerow.Core.System;
using Lexerow.Core.System.GenDef;
using Lexerow.Core.System.ScriptDef;
using Lexerow.Core.Utils;

namespace Lexerow.Core.ScriptCompile.Parse;

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
            if (scriptToken.Value.Equals(CoreInstr.InstrEndName, StringComparison.InvariantCultureIgnoreCase))
            {
                instrBase = new InstrEnd(scriptToken);
                return true;
            }

            // SelectFiles
            if (scriptToken.Value.Equals(CoreInstr.InstrSelectFiles, StringComparison.InvariantCultureIgnoreCase))
            {
                instrBase = new InstrSelectFiles(scriptToken);
                return true;
            }

            // OnExcel
            if (scriptToken.Value.Equals(CoreInstr.InstrOnExcel, StringComparison.InvariantCultureIgnoreCase))
            {
                instrBase = new InstrOnExcel(scriptToken);
                return true;
            }

            // OnSheet
            if (scriptToken.Value.Equals(CoreInstr.InstrOnSheet, StringComparison.InvariantCultureIgnoreCase))
            {
                instrBase = new InstrOnSheet(scriptToken);
                return true;
            }

            // ForEach
            if (scriptToken.Value.Equals(CoreInstr.InstrForEach, StringComparison.InvariantCultureIgnoreCase))
            {
                instrBase = new InstrForEach(scriptToken);
                return true;
            }

            // Special case: ForEachRow is allowed -> Row
            if (scriptToken.Value.Equals(CoreInstr.InstrForEachRow, StringComparison.InvariantCultureIgnoreCase))
            {
                instrBase = new InstrForEachRow(scriptToken);
                return true;
            }

            // Row
            if (scriptToken.Value.Equals(CoreInstr.InstrRow, StringComparison.InvariantCultureIgnoreCase))
            {
                instrBase = new InstrRow(scriptToken);
                return true;
            }

            // Next
            if (scriptToken.Value.Equals(CoreInstr.InstrNext, StringComparison.InvariantCultureIgnoreCase))
            {
                instrBase = new InstrNext(scriptToken);
                return true;
            }

            // Col
            if (scriptToken.Value.Equals(CoreInstr.InstrCol, StringComparison.InvariantCultureIgnoreCase))
            {
                instrBase = new InstrCol(scriptToken);
                return true;
            }

            // Cell
            if (scriptToken.Value.Equals(CoreInstr.InstrCell, StringComparison.InvariantCultureIgnoreCase))
            {
                instrBase = new InstrCell(scriptToken);
                return true;
            }

            // if
            if (scriptToken.Value.Equals(CoreInstr.InstrIf, StringComparison.InvariantCultureIgnoreCase))
            {
                instrBase = new InstrIf(scriptToken);
                return true;
            }

            // then
            if (scriptToken.Value.Equals(CoreInstr.InstrThen, StringComparison.InvariantCultureIgnoreCase))
            {
                instrBase = new InstrThen(scriptToken);
                return true;
            }

            // blank
            if (scriptToken.Value.Equals(CoreInstr.InstrBlank, StringComparison.InvariantCultureIgnoreCase))
            {
                instrBase = new InstrBlank(scriptToken);
                return true;
            }

            // null
            if (scriptToken.Value.Equals(CoreInstr.InstrNull, StringComparison.InvariantCultureIgnoreCase))
            {
                instrBase = new InstrNull(scriptToken);
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

            execResult.AddError(new ExecResultError(ErrorCode.ParserTokenNotExpected, scriptToken.Value));
            instrBase = null;
            return false;
        }

        //--script token is a string
        if (scriptToken.ScriptTokenType == ScriptTokenType.String)
        {
            // remove double quote
            string val = StringUtils.RemoveStartEndDoubleQuote(scriptToken.Value);
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

        execResult.AddError(new ExecResultError(ErrorCode.ParserTokenNotExpected, scriptToken.Value));
        instrBase = null;
        return false;
    }

    public static InstrSepComparison CreateSepComparison(ScriptToken scriptToken)
    {
        if (string.IsNullOrWhiteSpace(scriptToken.Value)) return null;
        return new InstrSepComparison(scriptToken);
    }
}