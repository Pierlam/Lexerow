using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;
using Lexerow.Core.System.InstrDef;
using Lexerow.Core.System.InstrDef.FuncCall;
using Lexerow.Core.System.ScriptCompile;
using Lexerow.Core.Utils;

namespace Lexerow.Core.ScriptCompile.Parse;

/// <summary>
/// Process the content of the stack.
/// Script End line reached.
/// </summary>
internal class ParserStackContentProcessor
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
    public static bool ScriptEndLineReached(IActivityLogger logger, Result result, List<InstrNameObject> listVar, int scriptLineNum, CompilStackInstr stackInstr, Program program)
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
            res = ProcessOnExcel(result, stackInstr, scriptLineNum, program.ListInstr, out isToken);
            if (!res) return false;
            if (isToken) return true;

            //--is previous instr on stack SetVar?  exp: a=12,  a=b, a=OpenExcel(), ...
            res = GatherSetVar(result, listVar, scriptLineNum, stackInstr, program.ListInstr, out isToken);
            if (!res) return false;
            if (isToken) continue;

            //--is previous instr on stack FirstRow value?
            res = ProcessFirstDataRowValue(result, listVar, scriptLineNum, stackInstr, program, out isToken);
            if (!res) return false;
            if (isToken) continue;

            //--is previous instr on stack ForEach?
            // TODO: never true, strange
            res = GatherForEachRow(result, listVar, scriptLineNum, stackInstr, program.ListInstr, out isToken);
            if (!res) return false;
            if (isToken) continue;

            //--is previous instr on stack Then?  exp: Then Inst
            res = GatherThen(result, listVar, scriptLineNum, stackInstr, program.ListInstr, out isToken);
            if (!res) return false;
            if (isToken) continue;

            //--is previous instr on stack If?
            res = GatherIfThen(result, stackInstr, scriptLineNum, program.ListInstr, out isToken);
            if (!res) return false;
            if (isToken) continue;

            //--is previous instr on stack End If?
            res = GatherEnd(logger, result, stackInstr, scriptLineNum, program.ListInstr, out isToken);
            if (!res) return false;
            if (isToken) continue;

            //--is it a fct call, exp: fct()
            res = PerformFctCall(logger, result, stackInstr, scriptLineNum, program.ListInstr, out isToken);
            if (!res) return false;
            // TODO: fct call can be in a Then or in a ForEachRow!
            if (isToken) continue;

            // get the last instr on the stack
            if (stackInstr.Count > 0)
                // case not managed, error or not yet implemented
                result.AddError(ErrorCode.ParserTokenNotExpected, stackInstr.Peek().FirstScriptToken());
            else
                // case not managed, error or not yet implemented
                result.AddError(ErrorCode.ParserTokenNotExpected, scriptLineNum.ToString());
            return false;
        }
    }

    /// <summary>
    /// Stack  IN: ...;  If; Then
    /// manage case Then without instr on the same line -> need a End If
    ///
    /// </summary>
    /// <param name="stackInstr"></param>
    /// <param name="isToken"></param>
    private static void ManageCaseThenWithoutInstrSameLine(CompilStackInstr stackInstr, out bool isToken)
    {
        isToken = false;

        if (stackInstr.Count == 0)
            return;

        // read the instr on top of the stack
        InstrBase instrBase = stackInstr.Peek();

        if (instrBase.InstrType != InstrType.Then)
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

    /// <summary>
    /// OnExcel instr found.
    /// Nothing more to do at this stage.
    /// </summary>
    /// <param name="result"></param>
    /// <param name="stackInstr"></param>
    /// <param name="sourceCodeLineIndex"></param>
    /// <param name="listInstrToExec"></param>
    /// <param name="isToken"></param>
    /// <returns></returns>
    private static bool ProcessOnExcel(Result result, CompilStackInstr stackInstr, int sourceCodeLineIndex, List<InstrBase> listInstrToExec, out bool isToken)
    {
        isToken = false;

        // read the instr on top of the satck
        var instrBase = stackInstr.Peek();

        // is it the OnExcel instr?
        InstrOnExcel instrOnExcel = instrBase as InstrOnExcel;
        if (instrOnExcel != null)
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
    private static bool GatherSetVar(Result result, List<InstrNameObject> listVar, int sourceCodeLineIndex, CompilStackInstr stackInstr, List<InstrBase> listInstrToExec, out bool isToken)
    {
        isToken = false;

        if (stackInstr.Count == 0)
            return true;

        // the instr before on top of the stack?  SetVar, instr
        var instrBefTop = stackInstr.ReadInstrBeforeTop();
        if (instrBefTop == null)
            // not a set var instr to finish
            return true;

        // the instr just before the top is a SetVar instr?
        if (instrBefTop.InstrType != InstrType.SetVar)
            // not a set var instr
            return true;

        isToken = true;

        // the left part of the instr is set previously
        InstrSetVar instrSetVar = instrBefTop as InstrSetVar;

        // the right part should be null
        if (instrSetVar.InstrRight != null)
        {
            result.AddError(ErrorCode.ParserVarWrongRightPart, instrSetVar.FirstScriptToken(), sourceCodeLineIndex.ToString());
            return false;
        }

        // get the second saved item
        InstrBase instrBase = stackInstr.Pop();

        // remove the first/oldest item, it's SetVar
        stackInstr.Pop();

        //--case a=12, A.Cell=12
        InstrValue instrValue = instrBase as InstrValue;
        if (instrValue != null)
        {
            instrSetVar.InstrRight = instrBase;
            instrSetVar.InstrLeft.ReturnType = InstrUtils.GetReturnType(instrValue);

            if (stackInstr.Count == 0)
                // instr SetVar not included in a Then instr or ForEachRow
                listInstrToExec.Add(instrSetVar);
            else
                stackInstr.Push(instrSetVar);
            return true;
        }

        //--case a=b
        InstrNameObject instrObjectName = instrBase as InstrNameObject;
        if (instrObjectName != null)
        {
            // can be a bool value: true or false, convert to a ValueBool 
            if(ValueUtils.IsBoolStrValue(instrObjectName.Name, out bool boolValue))
            {
                instrSetVar.InstrRight = new InstrValue(instrObjectName.FirstScriptToken(), boolValue);
                instrSetVar.InstrLeft.ReturnType = InstrReturnType.ValueBool;

                if (stackInstr.Count == 0)
                    // instr SetVar not included in a Then instr or ForEachRow
                    listInstrToExec.Add(instrSetVar);
                else
                    stackInstr.Push(instrSetVar);

                return true;
            }

            // Check that right var exists
            InstrNameObject instrObjectNameFound = listVar.FirstOrDefault(x => x.Name.Equals(instrObjectName.Name, StringComparison.InvariantCultureIgnoreCase));
            if (instrObjectNameFound != null)
            {
                // replace the left part by the saved var: same name and important: the return type is set!
                instrSetVar.InstrRight = instrObjectNameFound;
                instrSetVar.InstrLeft.ReturnType = instrObjectNameFound.ReturnType;

                if (stackInstr.Count == 0)
                    // instr SetVar not included in a Then instr or ForEachRow
                    listInstrToExec.Add(instrSetVar);
                else
                    stackInstr.Push(instrSetVar);

                return true;
            }
            result.AddError(ErrorCode.ParserVarWrongRightPart, instrBase.FirstScriptToken(), sourceCodeLineIndex.ToString());
            return false;
        }

        //--case x=blank
        InstrBlank instrBlank = instrBase as InstrBlank;
        if (instrBlank != null)
        {
            instrSetVar.InstrRight = instrBase;
            if (stackInstr.Count == 0)
                // instr SetVar not included in a Then instr or ForEachRow
                listInstrToExec.Add(instrSetVar);
            else
                stackInstr.Push(instrSetVar);
            return true;
        }

        //--case x=null
        InstrNull instrNull = instrBase as InstrNull;
        if (instrNull != null)
        {
            instrSetVar.InstrRight = instrBase;
            if (stackInstr.Count == 0)
                // instr SetVar not included in a Then instr or ForEachRow
                listInstrToExec.Add(instrSetVar);
            else
                stackInstr.Push(instrSetVar);
            return true;
        }

        //--case a=Fct(), apply the setVar
        if (instrBase.IsFunctionCall)
        {
            // check that the function return something to set to a var
            if (instrBase.ReturnType == InstrReturnType.Nothing)
            {
                result.AddError(ErrorCode.ParserVarWrongRightPart, instrBase.FirstScriptToken(), sourceCodeLineIndex.ToString());
                return false;
            }

            // check the fct call: params set and type ok?
            if (!InstrChecker.CheckFunctionCall(result, instrBase))
                return false;

            instrSetVar.InstrRight = instrBase;
            instrSetVar.InstrLeft.ReturnType = instrBase.ReturnType;

            if (stackInstr.Count == 0)
                // instr SetVar not included in a Then instr or ForEachRow
                listInstrToExec.Add(instrSetVar);
            else
                stackInstr.Push(instrSetVar);

            return true;
        }

        // other cases: unexpected so error
        result.AddError(ErrorCode.ParserVarWrongRightPart, instrBase.FirstScriptToken());
        return false;
    }

    /// <summary>
    /// is previous instr on stack FirstRow ?
    /// Stack  IN => ValueInt/VarName/FctCall; FirstRow, OnExcel
    /// Stack OUT => OnExcel
    ///
    /// OnExcel
    ///   [OnSheet sheetNum]
    ///   [FirstRow val/var]
    /// </summary>
    /// <param name="result"></param>
    /// <param name="scriptLineNum"></param>
    /// <param name="stackInstr"></param>
    /// <param name="isToken"></param>
    /// <returns></returns>
    private static bool ProcessFirstDataRowValue(Result result, List<InstrNameObject> listVar, int scriptLineNum, CompilStackInstr stackInstr, Program program, out bool isToken)
    {
        isToken = false;

        if (stackInstr.Count == 0)
            return true;

        // get the instr before on top of the stack
        var instrBefTop = stackInstr.ReadInstrBeforeTop();
        if (instrBefTop == null)
            return true;

        // the instr just before the top should be a constValue int
        if (instrBefTop.InstrType != InstrType.FirstRow)
            // not a then instr
            return true;

        isToken = true;

        // remove the top of the stack: value or varname or fctcall
        var instr = stackInstr.Pop();

        // then remove the FirstRow instr from the stack, as found before
        stackInstr.Pop();

        // the stack should contains one instr: OnExcel
        if (stackInstr.Count != 1)
        {
            // Instr OnExcel expected
            result.AddError(ErrorCode.ParserOnExcelExpected, scriptLineNum.ToString());
            return false;
        }

        // the next instr on the stack should be OnExcel
        InstrOnExcel instrOnExcel = stackInstr.Peek() as InstrOnExcel;
        if (instrOnExcel == null)
        {
            // Instr OnExcel expected
            result.AddError(ErrorCode.ParserOnExcelExpected, stackInstr.Peek().FirstScriptToken());
            return false;
        }

        // expected now/here -> check stage in OnExcel
        // TODO:

        //--check the instr value, should be an int
        if (!InstrUtils.GetIntFromInstrParser(result, program, instr, out bool valueSet, out int value))
            return false;
        if (valueSet)
        {
            // check the int value, should be >= 1
            if (value < 1)
            {
                result.AddError(ErrorCode.ParserValueIntWrong, instr.FirstScriptToken());
                return false;
            }
            // save the value into the current OnExcel sheet
            instrOnExcel.CurrOnSheet.InstrFirstDataRow = instr;
            return true;
        }

        //--the top instr on the stack is a fctcall?
        // TODO:

        result.AddError(ErrorCode.ParserCaseNotManaged, instr.FirstScriptToken());
        return false;
    }

    private static bool GatherForEachRow(Result result, List<InstrNameObject> listVar, int sourceCodeLineIndex, CompilStackInstr stackInstr, List<InstrBase> listInstrToExec, out bool isToken)
    {
        isToken = false;

        if (stackInstr.Count == 0)
            return true;

        // the instr before on top of the stack?  ForEachRow; instr
        var instrBefTop = stackInstr.ReadInstrBeforeTop();
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
        return true;
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
    /// <param name="result"></param>
    /// <param name="listVar"></param>
    /// <param name="sourceCodeLineIndex"></param>
    /// <param name="stackInstr"></param>
    /// <param name="listInstrToExec"></param>
    /// <param name="isToken"></param>
    /// <returns></returns>
    private static bool GatherThen(Result result, List<InstrNameObject> listVar, int sourceCodeLineIndex, CompilStackInstr stackInstr, List<InstrBase> listInstrToExec, out bool isToken)
    {
        isToken = false;

        if (stackInstr.Count == 0)
            return true;

        // the instr before on top of the stack?  Then; instr
        var instrBefTop = stackInstr.ReadInstrBeforeTop();
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
        if (instrBase is InstrEndIf)
        {
            instrThen.IsEndIfReached = true;
            isToken = true;
            return true;
        }

        // is this instr after Then in the script is in the same line?
        if (instrThen.FirstScriptToken().LineNum == instrBase.FirstScriptToken().LineNum)
            // very important to close properly the processing of If-Then instruction!
            instrThen.HasInstrAfterInSameLine = true;

        // save the then instr into the list
        instrThen.ListInstr.Add(instrBase);
        return true;
    }

    /// <summary>
    /// is the last instr on the stack is Then?
    /// Stack  IN: OnExcel, If, Then
    /// Stack OUT: OnExcel, IfThen
    /// </summary>
    /// <param name="result"></param>
    /// <param name="stkItems"></param>
    /// <param name="sourceCodeLineIndex"></param>
    /// <param name="listInstrToExec"></param>
    /// <param name="isToken"></param>
    /// <returns></returns>
    private static bool GatherIfThen(Result result, CompilStackInstr stackInstr, int sourceCodeLineIndex, List<InstrBase> listInstrToExec, out bool isToken)
    {
        isToken = false;

        if (stackInstr.Count == 0)
            return true;

        // the instr before on top of the stack?
        var instrBefTop = stackInstr.ReadInstrBeforeTop();
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
        InstrThen instrThen = instrBase as InstrThen;
        if (instrThen == null)
        {
            result.AddError(ErrorCode.ParserTokenThenExpected, instrBase.FirstScriptToken());
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
        {
            result.AddError(ErrorCode.ParserThenPartIsEmpty, instrBase.FirstScriptToken());
            return false;
        }

        // create instr IfThen based on instr If and Then
        InstrIfThenElse instrIfThenElse = new InstrIfThenElse(instrIf.FirstScriptToken());
        instrIfThenElse.InstrIf = instrIf;
        instrIfThenElse.InstrThen = instrThen;

        // OnExcel instr, add IfThen instr to the current OnSheet/ForEachRow
        bool res = InstrOnExcelBuilder.BuildIfThen(result, stackInstr, instrIfThenElse);
        if (!res) return false;

        isToken = true;
        return true;
    }

    /// <summary>
    /// is the last instr on the stack is End?
    /// Stack  IN: ...; End; If
    /// Stack OUT: ...; EndIf
    /// </summary>
    /// <param name="result"></param>
    /// <param name="stackInstr"></param>
    /// <param name="sourceCodeLineIndex"></param>
    /// <param name="listInstrToExec"></param>
    /// <param name="isToken"></param>
    /// <returns></returns>
    private static bool GatherEnd(IActivityLogger logger, Result result, CompilStackInstr stackInstr, int sourceCodeLineIndex, List<InstrBase> listInstrToExec, out bool isToken)
    {
        isToken = false;

        logger.LogCompilStart(ActivityLogLevel.Trace, "ParserStackContentProcessor.GatherEnd", "Stack.Count: " + stackInstr.Count);

        if (stackInstr.Count == 0)
            return true;

        // the instr before on top of the stack?
        var instrBefTop = stackInstr.ReadInstrBeforeTop();
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
        if (instrIf != null)
        {
            // remove the End If tokens from the stack
            stackInstr.Pop();
            stackInstr.Pop();

            InstrEndIf instrEndIf = new InstrEndIf(instrIf.FirstScriptToken());
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

        var error= result.AddNewError(ErrorCode.ParserTokenThenExpected, instrBase.FirstScriptToken());
        logger.LogCompilEndError(error, "ParserStackContentProcessor.GatherEnd",string.Empty);
        return false;
    }

    /// <summary>
    /// is it a fct call, WITHOUT setVar, exp: SelectFiles(), CopyHeader(), CopyRow(), ...
    /// </summary>
    /// <param name="result"></param>
    /// <param name="stackInstr"></param>
    /// <param name="listInstrToExec"></param>
    /// <param name="isToken"></param>
    /// <returns></returns>
    private static bool PerformFctCall(IActivityLogger logger, Result result, CompilStackInstr stackInstr, int sourceCodeLineIndex, List<InstrBase> listInstrToExec, out bool isToken)
    {
        isToken = false;
        logger.LogCompilStart(ActivityLogLevel.Trace, "ParserStackContentProcessor.PerformFctCall", "Stack.Count: " + stackInstr.Count);

        if (stackInstr.Count == 0)
            return true;

        // get the first/oldest item pushed in the stack
        var instrBase = stackInstr.Peek();
        

        // the first one is a SetVar instr?
        if (!instrBase.IsFunctionCall)
            // not a set var instr
            return true;

        // it's a fct call, the stack should contains only one item
        if (stackInstr.Count != 1)
        {
            result.AddError(ErrorCode.ParserTokenNotExpected, sourceCodeLineIndex.ToString());
            return false;
        }

        //--is it the fct SelectFiles ?
        InstrFuncCallSelectFiles instrFuncCallSelectFiles = instrBase as InstrFuncCallSelectFiles;
        if (instrFuncCallSelectFiles != null)
        {
            // OpenExcel result not used!
            result.AddError(ErrorCode.ParserFctResultNotSet, instrBase.FirstScriptToken(), sourceCodeLineIndex.ToString());
            return false;
        }

        //--is it the fct CopyHeader?
        InstrFuncCallCopyHeader instrFuncCallCopyHeader = instrBase as InstrFuncCallCopyHeader;
        if (instrFuncCallCopyHeader != null)
        {
            // ok! not a SetVar instr, but it's a fct call which is not expected without SetVar, so error
            isToken = true;
            stackInstr.Pop();
            listInstrToExec.Add(instrFuncCallCopyHeader);
            return true;
        }


        //--is it the fct CopyRow?
        InstrFuncCallCopyRow instrFuncCallCopyRow = instrBase as InstrFuncCallCopyRow;
        if (instrFuncCallCopyRow != null)
        {
            // ok 
            isToken = true;
            stackInstr.Pop();
            listInstrToExec.Add(instrFuncCallCopyRow);
            return true;
        }

        // other cases: unexpected -> error
        var error = result.AddNewError(ErrorCode.ParserTokenNotExpected, instrBase.FirstScriptToken(), sourceCodeLineIndex.ToString());
        logger.LogCompilEndError(error, "ParserStackContentProcessor.ProcessFctCall", "FuncCall without SetVar expected");
        return false;
    }

}