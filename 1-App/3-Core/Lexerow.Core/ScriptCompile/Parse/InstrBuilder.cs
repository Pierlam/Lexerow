using Lexerow.Core.System;
using Lexerow.Core.System.GenDef;
using Lexerow.Core.System.InstrDef;
using Lexerow.Core.System.InstrDef.InstrFuncDef;
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
    public static bool Build(Result result, ScriptToken scriptToken, out InstrBase instrBase)
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

            // OnExcel
            if (scriptToken.Value.Equals(CoreInstr.InstrOnExcel, StringComparison.InvariantCultureIgnoreCase))
            {
                instrBase = new InstrOnExcel(scriptToken);
                return true;
            }

            // OnSheet
            if (scriptToken.Value.Equals(CoreInstr.InstrOnSheet, StringComparison.InvariantCultureIgnoreCase))
            {
                InstrValue value= InstrUtils.CreateInstrValueInt(CoreInstr.FirstDataRowNum);
                instrBase = new InstrOnSheet(scriptToken, value);
                return true;
            }

            // FirstRow
            if (scriptToken.Value.Equals(CoreInstr.InstrFirstRow, StringComparison.InvariantCultureIgnoreCase))
            {
                instrBase = new InstrFirstRow(scriptToken);
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

            // FuncSelectFiles
            if (scriptToken.Value.Equals(CoreInstr.InstrFuncSelectFiles, StringComparison.InvariantCultureIgnoreCase))
            {
                instrBase = new InstrFuncSelectFiles(scriptToken);
                return true;
            }

            // FuncDate
            if (scriptToken.Value.Equals(CoreInstr.InstrFuncDate, StringComparison.InvariantCultureIgnoreCase))
            {
                instrBase = new InstrFuncDate(scriptToken);
                return true;
            }

            // if it is not a known keyword, it's an object name, can be: a variable or a user defined function
            instrBase = new InstrNameObject(scriptToken);
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
                instrBase = null;
                return true;
            }
            if (scriptToken.Value.Equals(".", StringComparison.InvariantCultureIgnoreCase))
            {
                instrBase = new InstrDot(scriptToken);
                return true;
            }
            if (scriptToken.Value.Equals(",", StringComparison.InvariantCultureIgnoreCase))
            {
                instrBase = new InstrComma(scriptToken);
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

            result.AddError(new ResultError(ErrorCode.ParserTokenNotExpected, scriptToken.Value));
            instrBase = null;
            return false;
        }

        //--script token is a string
        if (scriptToken.ScriptTokenType == ScriptTokenType.String)
        {
            // remove double quote
            string val = StringUtils.RemoveStartEndDoubleQuote(scriptToken.Value);
            instrBase = new InstrValue(scriptToken, val);
            return true;
        }

        //--script token is a int
        if (scriptToken.ScriptTokenType == ScriptTokenType.Integer)
        {
            instrBase = new InstrValue(scriptToken, scriptToken.ValueInt);
            return true;
        }

        //--script token is a double
        if (scriptToken.ScriptTokenType == ScriptTokenType.Double)
        {
            instrBase = new InstrValue(scriptToken, scriptToken.ValueDouble);
            return true;
        }

        result.AddError(new ResultError(ErrorCode.ParserTokenNotExpected, scriptToken.Value));
        instrBase = null;
        return false;
    }

    public static InstrSepComparison CreateSepComparison(ScriptToken scriptToken)
    {
        if (string.IsNullOrWhiteSpace(scriptToken.Value)) return null;
        return new InstrSepComparison(scriptToken);
    }
}