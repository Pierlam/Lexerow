using Lexerow.Core.System;
using Lexerow.Core.System.Compilator;
using NPOI.OpenXmlFormats.Spreadsheet;
using Org.BouncyCastle.Utilities.Collections;
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

        bool res;

        //--is it SetVar instr?  exp: a=12,  a=b, a=OpenExcel(), ...
        res = ProcessSetVar(execResult, listVar, sourceCodeLineIndex, stkItems, listInstrToExec, out bool isToken);
        if(!res)return false;
        if (isToken) return true;

        //--is if If instr?
        // TODO:

        //--is if Then instr?
        // TODO:

        //--is it a fct call, without setVar, exp: fct()
        res = ProcessFctCall(execResult, stkItems, sourceCodeLineIndex, listInstrToExec, out isToken);
        if (!res) return false;
        if (isToken) return true;


        // case not managed, error or not yet implemented
        execResult.AddError(ErrorCode.SyntaxAnalyzerTokenNotExpected, null, sourceCodeLineIndex.ToString());
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
    static bool ProcessSetVar(ExecResult execResult, List<InstrObjectName> listVar, int sourceCodeLineIndex, Stack<InstrBase> stkItems, List<InstrBase> listInstrToExec, out bool isToken)
    {
        isToken = false;

        if (stkItems.Count == 0)
            return true;

        // get the first/oldest item pushed in the stack
        var firstItem = stkItems.Last();

        // the first one is a SetVar instr?
        if (firstItem.InstrType != InstrType.SetVar)
            // not a set var instr
            return true;

        // setVar: 2 items expected in the stack
        if (stkItems.Count != 2)
        {
            execResult.AddError(ErrorCode.SyntaxAnalyzerSetVarWrongRightPart, null, sourceCodeLineIndex.ToString());
            return false;
        }
        isToken = true;

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

        //--case a=Fct(), apply the setVar
        if (instrBase.IsFunctionCall) 
        {
            // check that the function return something to set to a var
            if (instrBase.ReturnType == InstrFunctionReturnType.Nothing) 
            {
                execResult.AddError(ErrorCode.SyntaxAnalyzerSetVarWrongRightPart, instrBase.FirstScriptToken(), sourceCodeLineIndex.ToString());
                return false;
            }

            // check the fct call: params set and type ok?
            if (!InstrChecker.CheckFunctionCall(execResult, instrBase))
                return false;

            instrSetVar.InstrRight= instrBase;
            listInstrToExec.Add(instrSetVar);
            return true;
        }

        // other cases: unexpected -> error
        execResult.AddError(ErrorCode.SyntaxAnalyzerSetVarWrongRightPart, instrBase.FirstScriptToken(), sourceCodeLineIndex.ToString());
        return false;
    }

    /// <summary>
    /// is it a fct call, without setVar, exp: fct()
    /// </summary>
    /// <param name="execResult"></param>
    /// <param name="stkItems"></param>
    /// <param name="listInstrToExec"></param>
    /// <param name="isToken"></param>
    /// <returns></returns>
    static bool ProcessFctCall(ExecResult execResult, Stack<InstrBase> stkItems, int sourceCodeLineIndex, List<InstrBase> listInstrToExec, out bool isToken)
    {
        isToken = false;

        if (stkItems.Count == 0)
            return true;

        // get the first/oldest item pushed in the stack
        var instrBase = stkItems.Last();

        // the first one is a SetVar instr?
        if (!instrBase.IsFunctionCall)
            // not a set var instr
            return true;

        // it's a fct call, the stack should contains onyl one item
        if (stkItems.Count != 1)
        {
            execResult.AddError(ErrorCode.SyntaxAnalyzerTokenNotExpected, null, sourceCodeLineIndex.ToString());
            return false;
        }

        InstrOpenExcel instrOpenExcel = instrBase as  InstrOpenExcel;
        if (instrOpenExcel != null) 
        {
            // OpenExcel result not used, warning?
            execResult.AddError(ErrorCode.SyntaxAnalyzerFctResultNotSet, instrBase.FirstScriptToken(), sourceCodeLineIndex.ToString());
            return false;
        }


        // other cases: unexpected -> error
        execResult.AddError(ErrorCode.SyntaxAnalyzerTokenNotExpected, instrBase.FirstScriptToken(), sourceCodeLineIndex.ToString());
        return false;
    }
}
