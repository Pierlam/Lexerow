using Lexerow.Core.System;
using Lexerow.Core.System.GenDef;
using Lexerow.Core.System.InstrDef;
using Lexerow.Core.System.ScriptCompile;
using Lexerow.Core.System.ScriptDef;
using Lexerow.Core.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.ScriptCompile.Parse;


/// <summary>
/// Parse instruction colum cell function.
/// can be: A.Cell, A.Cell.BgColor, A.Cell.FgColor, 
/// excelOut.B.Cell, Sheet[1].B.Cell
/// The stack should contains items in the reverse order.
/// </summary>
public class InstrColCellFuncParser
{
    public static bool Parse(Result result, CompilStackInstr stackInstr,  ScriptToken scriptToken, out bool isToken)
    {
        isToken = false;

        //--case Cell 
        if (!ParserUtils.IsToken(CoreInstr.InstrCell, scriptToken) && !ParserUtils.IsToken(CoreInstr.InstrBgColor, scriptToken) && !ParserUtils.IsToken(CoreInstr.InstrFgColor, scriptToken))
            return true;

        // the stack contains nothing, error!
        if (stackInstr.Count == 0)
        {
            result.AddError(ErrorCode.ParserTokenNotExpected, scriptToken);
            return false;
        }

        //--case Cell 
        if (ParserUtils.IsToken(CoreInstr.InstrCell, scriptToken))
        {
            isToken = true;
            if (!ParseFuncCell(result, stackInstr, scriptToken))
                return false;

            // case excelOut.B.Cell, Sheet[1].B.Cell
            // TODO: get the instr from the stack -> if it's dot -> process it
            return true;
        }

        //--case BgColor
        if (ParserUtils.IsToken(CoreInstr.InstrBgColor, scriptToken))
        {
            // TODO: get the instr from the stack, should be a InstrColCellFunc
        }

        // error
        result.AddError(ErrorCode.ParserTokenNotExpected, scriptToken);
        return false;

    }

    /// <summary>
    /// Process case: A, ., Cell
    /// Stack> ., A
    /// The Cell token is not saved in the stack.
    /// </summary>
    /// <param name="result"></param>
    /// <param name="stackInstr"></param>
    /// <param name="scriptToken"></param>
    /// <returns></returns>
    public static bool ParseFuncCell(Result result, CompilStackInstr stackInstr, ScriptToken scriptToken)
    {
        // the stack should contains at least 3 tokens: A, ., Cell
        if (stackInstr.Count < 3)
        {
            result.AddError(ErrorCode.ParserTokenNotExpected, scriptToken);
            return false;
        }

        // the previous one: dot
        InstrBase instr= stackInstr.Pop();
        InstrDot instrDot = instr as InstrDot;
        if (instrDot == null)
        {
            result.AddError(ErrorCode.ParserTokenDotExpected, instr.FirstScriptToken());
            return false;
        }

        // the previous one: columne name, exp: A
        instr = stackInstr.Pop();
        InstrNameObject instrObjectName = instr as InstrNameObject;
        if (instrObjectName == null)
        {
            result.AddError(ErrorCode.ParserColAddressExpected, instr.FirstScriptToken());
            return false;
        }

        // get the colNum based on the col name
        int colNum = ExcelUtils.ColumnNameToNumber(instrObjectName.Name);
        if (colNum < 1)
        {
            result.AddError(ErrorCode.ParserColNumWrong, instrObjectName.FirstScriptToken());
            return false;
        }

        // ok create the Column address instr and push into the stack
        InstrColCellFunc instrColCellFunc = new InstrColCellFunc(instrObjectName.FirstScriptToken(), InstrColCellFuncType.Value, instrObjectName.Name, colNum);
        stackInstr.Push(instrColCellFunc);
        return true;
    }
}
