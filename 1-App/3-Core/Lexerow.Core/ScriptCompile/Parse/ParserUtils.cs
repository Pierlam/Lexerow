using Lexerow.Core.System;
using Lexerow.Core.System.InstrDef;
using Lexerow.Core.System.ScriptCompile;
using Lexerow.Core.System.ScriptDef;
using Lexerow.Core.Utils;

namespace Lexerow.Core.ScriptCompile.Parse;
internal class ParserUtils
{
    /// <summary>
    /// Is it instr column Cell function?
    /// exp: A.Cell, A.Cell.BgColor, VB.Cell.FgColor,...
    ///
    /// The stack should contains items in reverse order.
    /// the last token is on the top of stack: Cell or BgColor or FgColor.
    /// </summary>
    /// <param name="result"></param>
    /// <param name="stackInstr"></param>
    /// <param name="isInstr"></param>
    /// <returns></returns>
    public static bool ProcessInstrColCellFunc(Result result, CompilStackInstr stackInstr, ScriptToken scriptToken, out bool isInstr)
    {
        isInstr = false;
        if (stackInstr.Count == 0) return true;

        var lastInstr = stackInstr.Peek();
        bool res;

        //--is it Cell token?
        if (lastInstr.InstrType == InstrType.Cell)
        {
            isInstr = true;
            lastInstr = stackInstr.Pop();

            // still 2 saved instr are expected
            if (stackInstr.Count < 2)
            {
                result.AddError(ErrorCode.ParserTokenNotExpected, scriptToken);
                return false;
            }

            res = IsInstrColDot(result, stackInstr, out InstrNameObject instrObjectName);

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

        //--is it BgColor token?
        //if (lastInstr.InstrType == InstrType.BgColor)
        //{
        //    // TODO:
        //}

        //--is it FgColor token?
        // TODO:

        // not a Col.Cell instr
        isInstr = false;
        return true;
    }

    public static bool IsMathOperator(InstrBase instr)
    {
        if (instr is InstrCharPlus) return true;
        if (instr is InstrCharMinus) return true;
        if (instr is InstrCharMul) return true;
        if (instr is InstrCharDiv) return true;

        return false;
    }

    public static bool IsComparisonSeparator(ScriptToken currToken)
    {
        if (currToken.Value.Equals("=")) return true;
        if (currToken.Value.Equals(">")) return true;
        if (currToken.Value.Equals("<")) return true;
        if (currToken.Value.Equals("<>")) return true;
        if (currToken.Value.Equals(">=")) return true;
        if (currToken.Value.Equals("<=")) return true;
        return false;
    }

    public static bool IsValueString(ValueBase value)
    {
        if (value == null) return false;
        if (value.ValueType == System.ValueType.String) return true;

        return false;
    }

    static bool IsInstrColDot(Result result, CompilStackInstr stkInstr, out InstrNameObject instrObjectName)
    {
        instrObjectName = null;

        // so previous one should be the dot token
        var instrBaseDot = stkInstr.Pop();
        InstrDot instrDot = instrBaseDot as InstrDot;
        if (instrDot == null)
        {
            result.AddError(ErrorCode.ParserTokenDotExpected, instrBaseDot.FirstScriptToken());
            return false;
        }

        // the last saved should be column address, exp: A
        var instrBaseColAddr = stkInstr.Pop();
        instrObjectName = instrBaseColAddr as InstrNameObject;
        if (instrObjectName == null)
        {
            result.AddError(ErrorCode.ParserColAddressExpected, instrBaseDot.FirstScriptToken());
            return false;
        }

        return true;
    }
}