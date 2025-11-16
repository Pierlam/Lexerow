using Lexerow.Core.System;
using Lexerow.Core.System.ScriptCompile;
using Lexerow.Core.System.ScriptDef;

namespace Lexerow.Core.ScriptCompile.Parse;

internal class SetVarDecoder
{
    /// <summary>
    /// Manage case  equal char is found, as a set variable instruction.
    /// exp:
    ///   file=                 implemented
	///   A.Cell= 	            implemented
	///   A.Cell.BgColor =      TODO
	///   excelOut.B.Cell=      TODO
	///   Sheet[1].B.Cell=      TODO?
    /// </summary>
    /// <param name="execResult"></param>
    /// <param name="stackInstr"></param>
    /// <param name="scriptToken"></param>
    /// <param name="listExecTok"></param>
    /// <param name="isToken"></param>
    /// <returns></returns>
    public static bool ProcessSetVarEqualChar(ExecResult execResult, List<InstrObjectName> listVar, CompilStackInstr stackInstr, ScriptToken scriptToken, List<InstrBase> listExecTok, out bool isToken)
    {
        isToken = false;
        bool res;

        // is the script token the equal char?
        if (!scriptToken.Value.Equals("=", StringComparison.InvariantCultureIgnoreCase))
            // not the equal char, bye without error
            return true;

        // the stack contains nothing, strange  =blabla
        if (stackInstr.Count == 0)
        {
            execResult.AddError(ErrorCode.ParserTokenNotExpected, scriptToken);
            return false;
        }

        // is it var=  ?
        if (stackInstr.Count == 1)
        {
            res = ProcessVarName(execResult, listVar, stackInstr, scriptToken, out isToken);
            if (!isToken)
            {
                execResult.AddError(ErrorCode.ParserTokenNotExpected, scriptToken);
                return false;
            }

            return true;
        }

        // is it a SetVar instr or Comparison? just one instr -> SetVar: instr=
        if (stackInstr.Count > 1)
        {
            // looking in the stack for Then instr or If. If not found search: ForEach Row
            InstrBase instrResult = stackInstr.FindFirstFromTop(InstrType.If, InstrType.Then);
            if (instrResult != null && instrResult.InstrType == InstrType.If)
                // close to an If instr, it's a comparison, not a SetVar
                return true;
        }

        isToken = true;

        // is it varName= ?
        res = ProcessVarName(execResult, listVar, stackInstr, scriptToken, out isToken);
        if (!res) return false;
        if (isToken) return true;

        //--is the stack contains A.Cell expression?
        res = ParserUtils.ProcessInstrColCellFunc(execResult, stackInstr, scriptToken, out isToken);
        if (!res) return false;
        if (isToken)
        {
            // the Instr Col.Cell.Function is on the top the stack
            InstrBase instrColCellFunc = stackInstr.Pop();
            // now process the SetVar instr
            InstrSetVar instrSetVar = new InstrSetVar(instrColCellFunc.FirstScriptToken());
            instrSetVar.InstrLeft = instrColCellFunc;
            instrSetVar.ListScriptToken.Add(scriptToken);

            stackInstr.Push(instrSetVar);
            return true;
        }

        execResult.AddError(ErrorCode.ParserTokenNotExpected, stackInstr.Peek().FirstScriptToken());
        return false;
    }

    /// <summary>
    /// Is the stack contains: varName ?
    /// </summary>
    /// <param name="execResult"></param>
    /// <param name="listVar"></param>
    /// <param name="stkInstr"></param>
    /// <param name="scriptToken"></param>
    /// <param name="isToken"></param>
    /// <returns></returns>
    private static bool ProcessVarName(ExecResult execResult, List<InstrObjectName> listVar, CompilStackInstr stkInstr, ScriptToken scriptToken, out bool isToken)
    {
        isToken = false;

        // the stack contains one item which is an Object name, exp: file=
        if (stkInstr.Peek().InstrType != InstrType.ObjectName)
            return true;

        isToken = true;

        InstrObjectName instrObjectName = stkInstr.Peek() as InstrObjectName;
        InstrSetVar instrSetVar = new InstrSetVar(instrObjectName.FirstScriptToken());
        instrSetVar.InstrLeft = instrObjectName;

        instrSetVar.ListScriptToken.Add(scriptToken);

        // remove the objectName from the stack and push the SetVar instr
        stkInstr.Pop();
        stkInstr.Push(instrSetVar);

        // save the var
        listVar.Add(instrObjectName);

        return true;
    }
}