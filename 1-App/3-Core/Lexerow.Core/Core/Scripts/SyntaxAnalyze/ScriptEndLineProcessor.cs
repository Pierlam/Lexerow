using Lexerow.Core.System;
using Lexerow.Core.System.Compilator;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Scripts.SyntaxAnalyze;

/// <summary>
/// Script End line reached.
/// No more token to process.
/// So process the content of instr saved in the stack.
/// </summary>
internal class ScriptEndLineProcessor
{
    /// <summary>
    /// no more token in the current line tokens, process items saved in the stack.
    /// cases:
    /// a=12
    /// b=a
    /// if a=12
    /// if sheet.Cell(A,1)=12 
    /// </summary>
    /// <param name="stkItems"></param>
    /// <param name="token"></param>
    /// <param name="compiledScript"></param>
    /// <returns></returns>
    public static bool ScriptEndLineReached(ExecResult execResult, List<InstrObjectName> listVar, int sourceCodeLineIndex, Stack<InstrBase> stkItems, List<InstrBase> listInstrToExec)
    {
        // no more item in the stack
        if (stkItems.Count == 0) return true;

        //--2 items in the stack, is it SetVar instr?
        if (ProcessSetVar(execResult, listVar, sourceCodeLineIndex, stkItems, listInstrToExec))
            return true;

        //--case a=12
        // TODO:

        //--case a=b
        // TODO:

        //--case a=Fct()
        // TODO:


        // other cases: unexpected -> error
        // TODO:

        return false;
    }

    /// <summary>
    /// Process SetVar instruction.
    /// a=12
    /// b=a
    /// file=OpenExcel(x)     SetVar, OpenExcel
    /// sheet.Cell(A,1)=12 
    /// 
    /// </summary>
    /// <param name="stkItems"></param>
    /// <param name="compiledScript"></param>
    /// <returns></returns>
    static bool ProcessSetVar(ExecResult execResult, List<InstrObjectName> listVar, int sourceCodeLineIndex, Stack<InstrBase> stkItems, List<InstrBase> listInstrToExec)
    {
        if (stkItems.Count != 2)
        {
            execResult.AddError(ErrorCode.SyntaxAnalyzerSetVarWrongRightPart, null, sourceCodeLineIndex.ToString());
            return false;
        }

        // get the first/oldest item pushed in the stack
        var firstItem = stkItems.Last();

        // the first one is a SetVar instr?  can be an if instr
        if (firstItem.InstrType != InstrType.SetVar)
            // TODO: sortir une erreur et tracer
            return false;

        // the left part of the instr is set previously
        InstrSetVar instrSetVar = firstItem as InstrSetVar;

        // get the second saved item 
        InstrBase instrBase = stkItems.Pop();

        // remove the first/oldest item, it's SetVar
        stkItems.Pop();

        //--case a=12
        InstrConstValue instrConstValue = instrBase as InstrConstValue;
        if (instrConstValue != null) 
        {
            instrSetVar.InstrRight = instrBase;
            listInstrToExec.Add(instrSetVar);
            return true;
        }

        //--case a=b
        InstrObjectName instrObjectName = instrBase as InstrObjectName;
        if (instrObjectName != null) 
        {
            // Check that right var exists
            if(listVar.FirstOrDefault(x=>x.ObjectName.Equals(instrObjectName.ObjectName, StringComparison.InvariantCultureIgnoreCase))!=null)
            {
                instrSetVar.InstrRight = instrObjectName;
                listInstrToExec.Add(instrSetVar);
                return true;
            }
            execResult.AddError(ErrorCode.SyntaxAnalyzerSetVarWrongRightPart, instrBase.FirstScriptToken(), sourceCodeLineIndex.ToString());
            return false;
        }

        //--case a=Fct()
        if (instrBase.IsFunctionCall) 
        {
            // check that the function return something to set to a var
            if (instrBase.ReturnType == InstrFunctionReturnType.Nothing) 
            {
                execResult.AddError(ErrorCode.SyntaxAnalyzerSetVarWrongRightPart, instrBase.FirstScriptToken(), sourceCodeLineIndex.ToString());
                return false;
            }

            instrSetVar.InstrRight= instrBase;
            listInstrToExec.Add(instrSetVar);
            return true;
        }

        // other cases: unexpected -> error
        execResult.AddError(ErrorCode.SyntaxAnalyzerSetVarWrongRightPart, instrBase.FirstScriptToken(), sourceCodeLineIndex.ToString());
        return false;
    }

}
