using Lexerow.Core.System;
using Lexerow.Core.System.Compilator;
using Lexerow.Core.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Intrinsics.X86;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Core.Scripts;
internal class SyntaxAnalyserUtils
{
    /// <summary>
    /// Is it instr column Cell function?
    /// exp: A.Cell, A.Cell.BgColor, VB.Cell.FgColor,...
    /// 
    /// The stack should contains items in reverse order.
    /// the last token is on the top of stack: Cell or BgColor or FgColor.
    /// </summary>
    /// <param name="execResult"></param>
    /// <param name="stkInstr"></param>
    /// <param name="isInstr"></param>
    /// <returns></returns>
    public static bool ProcessInstrColCellFunc(ExecResult execResult, Stack<InstrBase> stkInstr, ScriptToken scriptToken, out bool isInstr)
    {
        var lastInstr = stkInstr.Peek();
        bool res;

        //--is it Cell token?
        if (lastInstr.InstrType == InstrType.Cell)
        {
            isInstr = true;
            lastInstr = stkInstr.Pop();

            // still 2 saved instr are expected
            if (stkInstr.Count < 2)
            {
                execResult.AddError(ErrorCode.SyntaxAnalyzerTokenNotExpected, scriptToken);
                return false;
            }

            res= IsInstrColDot(execResult, stkInstr, out InstrObjectName instrObjectName);

            // get the colNum based on the col name
            int colNum = ExcelUtils.ColumnNameToNumber(instrObjectName.ObjectName);
            if(colNum<1)
            {
                execResult.AddError(ErrorCode.SyntaxAnalyzerColNumWrong, instrObjectName.FirstScriptToken());
                return false;
            }

            // ok create the Column address instr and push into the stack
            InstrColCellFunc instrColCellFunc = new InstrColCellFunc(instrObjectName.FirstScriptToken(), InstrColCellFuncType.Value, instrObjectName.ObjectName, colNum);
            stkInstr.Push(instrColCellFunc);
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

    /// <summary>
    /// Looking for an instr in the stack, starting from the top.
    /// </summary>
    /// <param name="stkInstr"></param>
    /// <param name="type"></param>
    /// <returns></returns>
    public static InstrBase FindFirstFromTop(Stack<InstrBase> stkInstr, InstrType type, InstrType type2)
    {
        foreach (var instr in stkInstr)
        {
            if(instr.InstrType == type)
                return instr;
            if (instr.InstrType == type2)
                return instr;
        }
        return null;
    }

    /// <summary>
    /// Get the inst just before the instr on top of the stack.
    /// </summary>
    /// <param name="stkInstr"></param>
    /// <returns></returns>
    public static InstrBase GetBeforeTop(Stack<InstrBase> stkInstr)
    {
        // need to ahve 2 isntr on the stack
        if (stkInstr.Count < 2) return null;
        return stkInstr.ElementAt(1);
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
        if(value == null) return false;
        if(value.ValueType== System.ValueType.String) return true;

        return false;
    }

    static bool IsInstrColDot(ExecResult execResult, Stack<InstrBase> stkInstr, out InstrObjectName instrObjectName)
    {
        instrObjectName = null;

        // so previous one should be the dot token
        var instrBaseDot = stkInstr.Pop();
        InstrDot instrDot = instrBaseDot as InstrDot;
        if (instrDot == null)
        {
            execResult.AddError(ErrorCode.SyntaxAnalyzerTokenDotExpected, instrBaseDot.FirstScriptToken());
            return false;
        }

        // the last saved should be column address, exp: A
        var instrBaseColAddr = stkInstr.Pop();
        instrObjectName = instrBaseColAddr as InstrObjectName;
        if (instrObjectName == null)
        {
            execResult.AddError(ErrorCode.SyntaxAnalyzerColAddressExpected, instrBaseDot.FirstScriptToken());
            return false;
        }

        return true;
    }

}
