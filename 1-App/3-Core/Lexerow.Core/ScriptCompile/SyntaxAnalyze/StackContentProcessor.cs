using Lexerow.Core.System;
using Lexerow.Core.System.Compilator;

namespace Lexerow.Core.ScriptCompile.SyntaxAnalyze;

/// <summary>
/// Process the content of the stack.
/// Script End line reached.
/// </summary>
internal class StackContentProcessor
{
    /// <summary>
    /// no more token in the current line tokens, process items saved in the stack.
    /// cases:
    /// file=OpenExcel(x) -> SetVar, OpenExcel
    /// a=12
    /// b=a
    /// Then A.CellA=12   -> ..., InstrThen, InstrColCellFunc, SetVar, 12 
    /// Then a=12
    /// If A.Cell=12
    /// If a=12
    /// </summary>
    /// <param name="stackInstr"></param>
    /// <param name="token"></param>
    /// <param name="compiledScript"></param>
    /// <returns></returns>
    public static bool ScriptEndLineReached(ExecResult execResult, List<InstrObjectName> listVar, int sourceCodeLineIndex, CompilStackInstr stackInstr, List<InstrBase> listInstrToExec)
    {
        bool res, isToken;

        while (true)
        {
            // no more item in the stack
            if (stackInstr.Count == 0) return true;

            //--manage case Then without instr on the same line -> need a End If
            ManageCaseThenWithoutInstrSameLine(stackInstr, out isToken);
            // no more to do in this fct, have to proceed next script lines
            if (isToken) return true;

            //--is it OnExcel instr ? 
            res = ProcessOnExcel(execResult, stackInstr, sourceCodeLineIndex, listInstrToExec, out isToken);
            if (!res) return false;
            if (isToken) return true;

            //--is previous instr on stack SetVar?  exp: a=12,  a=b, a=OpenExcel(), ...
            res = GatherSetVar(execResult, listVar, sourceCodeLineIndex, stackInstr, listInstrToExec, out isToken);
            if (!res) return false;
            if (isToken) continue;

            //--is previous instr on stack Then?  exp: Then Inst
            res = GatherThen(execResult, listVar, sourceCodeLineIndex, stackInstr, listInstrToExec, out isToken);
            if (!res) return false;
            if (isToken) continue;

            //--is previous instr on stack ForEach?  
            // TODO: never true, strange
            res = GatherForEachRow(execResult, listVar, sourceCodeLineIndex, stackInstr, listInstrToExec, out isToken);
            if (!res) return false;
            if (isToken) continue;

            //--is previous instr on stack If?  
            res = GatherIfThen(execResult, stackInstr, sourceCodeLineIndex, listInstrToExec, out isToken);
            if (!res) return false;
            if (isToken) continue;

            //--is previous instr on stack End If?  
            res = GatherEnd(execResult, stackInstr, sourceCodeLineIndex, listInstrToExec, out isToken);
            if (!res) return false;
            if (isToken) continue;

            //--is on top of the satck instr on Stack is Then?
            //res = ProcessThen(execResult, stackInstr, sourceCodeLineIndex, listInstrToExec, out isToken);
            //if (!res) return false;
            //if (isToken) continue;

            //--is it a fct call, exp: fct()
            res = ProcessFctCall(execResult, stackInstr, sourceCodeLineIndex, listInstrToExec, out isToken);
            if (!res) return false;
            // TODO: fct call can be in a Then or in a ForEachRow!
            if (isToken) continue;

            // case not managed, error or not yet implemented
            execResult.AddError(ErrorCode.ParserTokenNotExpected, null, sourceCodeLineIndex.ToString());
            return false;
        }
    }

    /// <summary>
    /// OnExcel instr found.
    /// Nothing more to do at this stage.
    /// </summary>
    /// <param name="execResult"></param>
    /// <param name="stackInstr"></param>
    /// <param name="sourceCodeLineIndex"></param>
    /// <param name="listInstrToExec"></param>
    /// <param name="isToken"></param>
    /// <returns></returns>
    static bool ProcessOnExcel(ExecResult execResult, CompilStackInstr stackInstr, int sourceCodeLineIndex, List<InstrBase> listInstrToExec, out bool isToken)
    {
        isToken = false;

        // read the instr on top of the satck
        var instrBase = stackInstr.Peek();

        // is it the OnExcel instr?
        InstrOnExcel instrOnExcel = instrBase as InstrOnExcel;
        if(instrOnExcel != null)  
            isToken = true;

        return true;
    }

    /// <summary>
    /// Process SetVar instruction.
    /// Finish the processing of a SetVar instr.
    /// Stack IN: SetVar, Instr   ->Instr which is right part.
    /// 
    /// -SetVar stand-alone:
    ///   file=OpenExcel(x) -> SetVar, OpenExcel
    ///   a=12   -> 
    ///   b=a    -> 
    ///
    /// -SetVar in a Then instr:
    ///   Then A.CellA=12   -> ..., InstrThen, SetVar:Left=InstrColCellFunc, 12 
    /// 
    /// -SetVar in ForEach Row:
    ///   ForEach Row
    ///      A.CellA=12
    /// </summary>
    /// <param name="stackInstr"></param>
    /// <param name="compiledScript"></param>
    /// <returns></returns>
    static bool GatherSetVar(ExecResult execResult, List<InstrObjectName> listVar, int sourceCodeLineIndex, CompilStackInstr stackInstr, List<InstrBase> listInstrToExec, out bool isToken)
    {
        isToken = false;

        if (stackInstr.Count == 0)
            return true;

        // the instr before on top of the stack?  SetVar, instr
        var instrBefTop= stackInstr.GetBeforeTop();
        if(instrBefTop==null)
            // not a set var instr to finish
            return true;

        // the instr just before the topt is a SetVar instr?
        if (instrBefTop.InstrType != InstrType.SetVar)
            // not a set var instr
            return true;

        isToken = true;

        // the left part of the instr is set previously
        InstrSetVar instrSetVar = instrBefTop as InstrSetVar;

        // the right part should be null
        if (instrSetVar.InstrRight != null)
            // TODO: to manage in a better way!
            throw new Exception("SetVar: Right part should not yet set!");

        // get the second saved item 
        InstrBase instrBase = stackInstr.Pop();

        // remove the first/oldest item, it's SetVar
        stackInstr.Pop();

        //--case a=12, A.Cell=12
        InstrConstValue instrConstValue = instrBase as InstrConstValue;
        if (instrConstValue != null) 
        {
            instrSetVar.InstrRight = instrBase;
            if(stackInstr.Count==0)
                // instr SetVar not included in a Then instr or ForEachRow
                listInstrToExec.Add(instrSetVar);
            else
                stackInstr.Push(instrSetVar);
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
                if (stackInstr.Count == 0)
                    // instr SetVar not included in a Then instr or ForEachRow
                    listInstrToExec.Add(instrSetVar);
                else
                    stackInstr.Push(instrSetVar);

                return true;
            }
            execResult.AddError(ErrorCode.ParserSetVarWrongRightPart, instrBase.FirstScriptToken(), sourceCodeLineIndex.ToString());
            return false;
        }

        //--case a=Fct(), apply the setVar
        if (instrBase.IsFunctionCall) 
        {
            // check that the function return something to set to a var
            if (instrBase.ReturnType == InstrFunctionReturnType.Nothing) 
            {
                execResult.AddError(ErrorCode.ParserSetVarWrongRightPart, instrBase.FirstScriptToken(), sourceCodeLineIndex.ToString());
                return false;
            }

            // check the fct call: params set and type ok?
            if (!InstrChecker.CheckFunctionCall(execResult, instrBase))
                return false;

            instrSetVar.InstrRight= instrBase;
            if (stackInstr.Count == 0)
                // instr SetVar not included in a Then instr or ForEachRow
                listInstrToExec.Add(instrSetVar);
            else
                stackInstr.Push(instrSetVar);

            return true;
        }

        // other cases: unexpected so error
        execResult.AddError(ErrorCode.ParserSetVarWrongRightPart, instrBase.FirstScriptToken());
        return false;
    }


    /// <summary>
    /// Gather then instr.
    /// Stack  IN: ...;  If; Then; Instr 
    /// Stack OUT: ...;  If; Then
    /// 
    /// Oter case:
    /// Stack  IN: ...;  If; Then; EndIf
    /// Stack OUT: ...;  If; Then
    /// </summary>
    /// <param name="execResult"></param>
    /// <param name="listVar"></param>
    /// <param name="sourceCodeLineIndex"></param>
    /// <param name="stackInstr"></param>
    /// <param name="listInstrToExec"></param>
    /// <param name="isToken"></param>
    /// <returns></returns>
    static bool GatherThen(ExecResult execResult, List<InstrObjectName> listVar, int sourceCodeLineIndex, CompilStackInstr stackInstr, List<InstrBase> listInstrToExec, out bool isToken)
    {
        isToken = false;

        if (stackInstr.Count == 0)
            return true;

        // the instr before on top of the stack?  Then; instr
        var instrBefTop = stackInstr.GetBeforeTop();
        if (instrBefTop == null)
            return true;

        // the instr just before the topt is a then instr?
        if (instrBefTop.InstrType != InstrType.Then)
            // not a then instr
            return true;

        isToken = true;

        InstrThen instrThen = instrBefTop as InstrThen;

        // extract the top instr
        InstrBase instrBase = stackInstr.Pop();

        // top instr is EndIf?
        if(instrBase is InstrEndIf)
        {
            instrThen.IsEndIfReached = true;
            isToken = true;
            return true;
        }

        // is this instr after Then in the script is in the same line?
        if (instrThen.FirstScriptToken().LineNum == instrBase.FirstScriptToken().LineNum)
            instrThen.HasInstrAfterInSameLine = true;

            // save the then instr into the list
        instrThen.ListInstr.Add(instrBase);
        return true;
    }

    /// <summary>
    /// Stack  IN: ...;  If; Then
    /// manage case Then without instr on the same line -> need a End If
    ///
    /// </summary>
    /// <param name="stackInstr"></param>
    /// <param name="isToken"></param>
    static void ManageCaseThenWithoutInstrSameLine(CompilStackInstr stackInstr, out bool isToken)
    {
        isToken = false;

        if (stackInstr.Count == 0)
            return;

        // read the instr on top of the stack
        InstrBase instrBase = stackInstr.Peek();

        if(instrBase.InstrType != InstrType.Then)
            return;

        InstrThen instrThen = instrBase as InstrThen;

        if (instrThen.HasInstrAfterInSameLine) 
            // not the case there is an instr after the token then on the same script line -> no EndIf expected
            return;

        if (instrThen.IsEndIfReached)
            return;

        // no instr after then, so expect an End If instr
        isToken = true;
    }

    static bool GatherForEachRow(ExecResult execResult, List<InstrObjectName> listVar, int sourceCodeLineIndex, CompilStackInstr stackInstr, List<InstrBase> listInstrToExec, out bool isToken)
    {
        isToken = false;

        if (stackInstr.Count == 0)
            return true;

        // the instr before on top of the stack?  ForEachRow; instr
        var instrBefTop = stackInstr.GetBeforeTop();
        if (instrBefTop == null)
            return true;

        // the instr just before the top is a then instr?
        if (instrBefTop.InstrType != InstrType.ForEach)
            // not a then instr
            return true;

        //XXX-TODO: never come here, strange

        isToken = true;

        InstrForEach instrForEach = instrBefTop as InstrForEach;

        // extract the top instr
        InstrBase instrBase = stackInstr.Pop();


        // remove the next which then instr
        stackInstr.Pop();

        //instrForEach.ListInstr.Add(instrBase);
        return true;

    }

    /// <summary>
    /// is the last instr on the stack is Then?
    /// Stack  IN: OnExcel, If, Then
    /// Stack OUT: OnExcel, IfThen
    /// </summary>
    /// <param name="execResult"></param>
    /// <param name="stkItems"></param>
    /// <param name="sourceCodeLineIndex"></param>
    /// <param name="listInstrToExec"></param>
    /// <param name="isToken"></param>
    /// <returns></returns>
    static bool GatherIfThen(ExecResult execResult, CompilStackInstr stackInstr, int sourceCodeLineIndex, List<InstrBase> listInstrToExec, out bool isToken)
    {
        isToken = false;

        if (stackInstr.Count == 0)
            return true;

        // the instr before on top of the stack?  
        var instrBefTop = stackInstr.GetBeforeTop();
        if (instrBefTop == null)
            return true;

        // the instr just before the topt is a then instr?
        if (instrBefTop.InstrType != InstrType.If)
            // not a then If
            return true;

        isToken = true;

        InstrIf instrIf = instrBefTop as InstrIf;

        // extract the top instr, should be Then
        InstrBase instrBase = stackInstr.Peek();
        InstrThen instrThen= instrBase as InstrThen;
        if(instrThen==null)
        {
            execResult.AddError(ErrorCode.ParserTokenThenExpected, instrBase.FirstScriptToken());
            return false;
        }

        // yes the instr If-Then expect to have the EndIf instr?
        //if (instrThen.IsEndIfInstrExpected)
        //    return true;

        // remove both If and Then instr from the top of the stack
        stackInstr.Pop();
        stackInstr.Pop();

        // check that there at leat one instr in the then part
        if (instrThen.ListInstr.Count == 0)
            // todo: manage it better!
            throw new Exception("No instr in Then!");

        // create instr IfThen based on instr If and Then
        InstrIfThenElse instrIfThenElse = new InstrIfThenElse(instrIf.FirstScriptToken());
        instrIfThenElse.InstrIf = instrIf;
        instrIfThenElse.InstrThen = instrThen;

        // OnExcel instr, add IfThen instr to the current OnSheet/ForEachRow        
        bool res = InstrOnExcelBuilder.BuildIfThen(execResult, stackInstr, instrIfThenElse);
        if (!res) return false;

        isToken = true;
        return true;
    }

    /// <summary>
    /// is the last instr on the stack is End?
    /// Stack  IN: ...; End; If
    /// Stack OUT: ...; EndIf
    /// </summary>
    /// <param name="execResult"></param>
    /// <param name="stackInstr"></param>
    /// <param name="sourceCodeLineIndex"></param>
    /// <param name="listInstrToExec"></param>
    /// <param name="isToken"></param>
    /// <returns></returns>
    static bool GatherEnd(ExecResult execResult, CompilStackInstr stackInstr, int sourceCodeLineIndex, List<InstrBase> listInstrToExec, out bool isToken)
    {
        isToken = false;

        if (stackInstr.Count == 0)
            return true;

        // the instr before on top of the stack?  
        var instrBefTop = stackInstr.GetBeforeTop();
        if (instrBefTop == null)
            return true;

        // the instr just before the topt is a End instr?
        if (instrBefTop.InstrType != InstrType.End)
            // not a then End
            return true;

        isToken = true;

        InstrEnd instrEnd = instrBefTop as InstrEnd;

        InstrBase instrBase = stackInstr.Peek();

        // --should be If or OnExcel
        InstrIf instrIf = instrBase as InstrIf;
        if(instrIf!=null)
        {
            // remove the End If tokens from the stack
            stackInstr.Pop();
            stackInstr.Pop();

            InstrEndIf instrEndIf= new InstrEndIf(instrIf.FirstScriptToken());
            stackInstr.Push(instrEndIf);
            return true;
        }

        // --should be OnExcel
        InstrOnExcel instrOnExcel = instrBase as InstrOnExcel;
        if (instrOnExcel != null)
        {
            // remove the End If tokens from the stack
            stackInstr.Pop();
            stackInstr.Pop();

            InstrEndOnExcel instrEndOnExcel = new InstrEndOnExcel(instrOnExcel.FirstScriptToken());
            stackInstr.Push(instrEndOnExcel);
            return true;
        }

        execResult.AddError(ErrorCode.ParserTokenThenExpected, instrBase.FirstScriptToken());
        return false;
    }

    /// <summary>
    /// is the last instr on the stack is Then?
    /// Stack IN: OnExcel, If, Then
    /// </summary>
    /// <param name="execResult"></param>
    /// <param name="stkItems"></param>
    /// <param name="sourceCodeLineIndex"></param>
    /// <param name="listInstrToExec"></param>
    /// <param name="isToken"></param>
    /// <returns></returns>
    static bool ProcessThen(ExecResult execResult, CompilStackInstr stackInstr, int sourceCodeLineIndex, List<InstrBase> listInstrToExec, out bool isToken)
    {
        isToken = false;

        // stack is empty, bye
        if (stackInstr.Count == 0) return true;

        // the instr on the top of the stack is Then?
        InstrThen instrThen = stackInstr.Peek() as InstrThen;
        if(instrThen==null) 
            // not the Then instr, bye
            return true;

        // yes the instr If-Then expect to have the EndIf instr?
        //if (instrThen.IsEndIfInstrExpected) 
        //    return true;

        // check that there at leat one instr in the then part
        if (instrThen.ListInstr.Count == 0)
            // todo: manage it better!
            throw new Exception("No instr in Then!");

        // remove the instr Then from the stack
        stackInstr.Pop();

        // the stack should contains the If instr
        if (stackInstr.Count == 0)
        {
            execResult.AddError(ErrorCode.ParserTokenIfExpected, instrThen.FirstScriptToken());
            return false;
        }

        // the previous one should be the instr if
        InstrIf instrIf = stackInstr.Pop() as InstrIf;
        if (instrIf==null)
        {
            execResult.AddError(ErrorCode.ParserTokenIfExpected, instrThen.FirstScriptToken());
            return false;
        }

        // create instr IfThen based on instr If and Then
        InstrIfThenElse instrIfThenElse = new InstrIfThenElse(instrIf.FirstScriptToken());
        instrIfThenElse.InstrIf = instrIf;
        instrIfThenElse.InstrThen = instrThen;

        // the stack should contains the If instr
        if (stackInstr.Count == 0)
        {
            execResult.AddError(ErrorCode.ParserTokenIfExpected, instrThen.FirstScriptToken());
            return false;
        }

        // OnExcel instr, add IfThen instr to the current OnSheet/ForEachRow        
        bool res= InstrOnExcelBuilder.BuildIfThen(execResult, stackInstr, instrIfThenElse);
        if (!res) return false;

        isToken = true;
        return true;
    }

    /// <summary>
    /// is it a fct call, without setVar, exp: fct()
    /// </summary>
    /// <param name="execResult"></param>
    /// <param name="stackInstr"></param>
    /// <param name="listInstrToExec"></param>
    /// <param name="isToken"></param>
    /// <returns></returns>
    static bool ProcessFctCall(ExecResult execResult, CompilStackInstr stackInstr, int sourceCodeLineIndex, List<InstrBase> listInstrToExec, out bool isToken)
    {
        isToken = false;

        if (stackInstr.Count == 0)
            return true;

        // get the first/oldest item pushed in the stack
        var instrBase = stackInstr.Peek();

        // the first one is a SetVar instr?
        if (!instrBase.IsFunctionCall)
            // not a set var instr
            return true;

        // it's a fct call, the stack should contains onyl one item
        if (stackInstr.Count != 1)
        {
            execResult.AddError(ErrorCode.ParserTokenNotExpected, null, sourceCodeLineIndex.ToString());
            return false;
        }

        //--is it the fct XXX ?
        // TODO:

        //--is it the fct InstrOpenExcel ?
        InstrSelectFiles instrOpenExcel = instrBase as InstrSelectFiles;
        if (instrOpenExcel != null)
        {
            // OpenExcel result not used!
            execResult.AddError(ErrorCode.ParserFctResultNotSet, instrBase.FirstScriptToken(), sourceCodeLineIndex.ToString());
            return false;
        }

        // other cases: unexpected -> error
        execResult.AddError(ErrorCode.ParserTokenNotExpected, instrBase.FirstScriptToken(), sourceCodeLineIndex.ToString());
        return false;
    }


}
