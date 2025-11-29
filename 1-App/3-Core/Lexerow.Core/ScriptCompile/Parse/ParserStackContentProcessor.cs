using Lexerow.Core.System;
using Lexerow.Core.System.InstrDef;
using Lexerow.Core.System.ScriptCompile;
using Lexerow.Core.Utils;
using NPOI.Util;
using Org.BouncyCastle.Utilities.Collections;
using System.Collections.Generic;

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
    public static bool ScriptEndLineReached(Result result, List<InstrNameObject> listVar, int scriptLineNum, CompilStackInstr stackInstr, Program program)
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
            res = GatherEnd(result, stackInstr, scriptLineNum, program.ListInstr, out isToken);
            if (!res) return false;
            if (isToken) continue;

            //--is it a fct call, exp: fct()
            res = ProcessFctCall(result, stackInstr, scriptLineNum, program.ListInstr, out isToken);
            if (!res) return false;
            // TODO: fct call can be in a Then or in a ForEachRow!
            if (isToken) continue;

            // get the last instr on the stack
            if(stackInstr.Count>0)
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

        // the instr just before the topt is a SetVar instr?
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
            // Check that right var exists
            if (listVar.FirstOrDefault(x => x.Name.Equals(instrObjectName.Name, StringComparison.InvariantCultureIgnoreCase)) != null)
            {
                instrSetVar.InstrRight = instrObjectName;
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
            if (instrBase.ReturnType == InstrFunctionReturnType.Nothing)
            {
                result.AddError(ErrorCode.ParserVarWrongRightPart, instrBase.FirstScriptToken(), sourceCodeLineIndex.ToString());
                return false;
            }

            // check the fct call: params set and type ok?
            if (!InstrChecker.CheckFunctionCall(result, instrBase))
                return false;

            instrSetVar.InstrRight = instrBase;
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
        if(stackInstr.Count != 1)
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
        if(!InstrUtils.CheckInstrInt(result, program, instr, out bool valueSet, out int value))
            return false;
        if(valueSet)
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

        ////--the top instr on the stack is a value string?
        //InstrValue instrValue = instr as InstrValue;
        //if (instrValue != null)
        //{
        //    if (!ProcessFirstDataRowValueValue(result, scriptLineNum, instrOnExcel, instrValue))
        //        return false;
        //    instrOnExcel.CurrOnSheet.InstrFirstDataRow = instrValue;
        //    return true;
        //}
        ////--the top instr on the stack is a varname?
        //InstrObjectName instrObjectName = instr as InstrObjectName;
        //if (instrObjectName != null)
        //{
        //    if (!ProcessFirstDataRowValueVar(result, listVar, scriptLineNum, program, instrOnExcel, instrObjectName))
        //        return false;
        //    instrOnExcel.CurrOnSheet.InstrFirstDataRow = instrObjectName;
        //    return true;
        //}

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
    private static bool GatherEnd(Result result, CompilStackInstr stackInstr, int sourceCodeLineIndex, List<InstrBase> listInstrToExec, out bool isToken)
    {
        isToken = false;

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

        result.AddError(ErrorCode.ParserTokenThenExpected, instrBase.FirstScriptToken());
        return false;
    }

    /// <summary>
    /// is it a fct call, without setVar, exp: fct()
    /// </summary>
    /// <param name="result"></param>
    /// <param name="stackInstr"></param>
    /// <param name="listInstrToExec"></param>
    /// <param name="isToken"></param>
    /// <returns></returns>
    private static bool ProcessFctCall(Result result, CompilStackInstr stackInstr, int sourceCodeLineIndex, List<InstrBase> listInstrToExec, out bool isToken)
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
            result.AddError(ErrorCode.ParserTokenNotExpected, sourceCodeLineIndex.ToString());
            return false;
        }

        //--is it the fct XXX ?
        // TODO:

        //--is it the fct SelectFiles ?
        InstrFuncSelectFiles instrOpenExcel = instrBase as InstrFuncSelectFiles;
        if (instrOpenExcel != null)
        {
            // OpenExcel result not used!
            result.AddError(ErrorCode.ParserFctResultNotSet, instrBase.FirstScriptToken(), sourceCodeLineIndex.ToString());
            return false;
        }

        // other cases: unexpected -> error
        result.AddError(ErrorCode.ParserTokenNotExpected, instrBase.FirstScriptToken(), sourceCodeLineIndex.ToString());
        return false;
    }

    //private static bool ProcessFirstDataRowValueValue(Result result, int scriptLineNum, InstrOnExcel instrOnExcel, InstrValue instrValue)
    //{
    //    if (!InstrUtils.GetValueIntFromInstrValue(result, instrValue, scriptLineNum, out int val))
    //        return false;

    //    // check the int value, should be >= 1
    //    if (val < 1)
    //    {
    //        result.AddError(ErrorCode.ParserValueIntWrong, instrValue.FirstScriptToken());
    //        return false;
    //    }
    //    // save the value into the current OnExcel sheet
    //    //instrOnExcel.CurrOnSheet.InstrFirstDataRow = instrValue;
    //    return true;
    //}

    //// should be a int value, constraint: >1
    //private static bool ProcessFirstDataRowValueVar(Result result, List<InstrObjectName> listVar, int scriptLineNum, Program program, InstrOnExcel instrOnExcel, InstrObjectName instrObjectName)
    //{
    //    // the varname should be defined
    //    // TODO: really needed?? program.FindLastVarSet() do the job!
    //    //var instrObjectNameVar = listVar.FirstOrDefault(v => v.MatchName(instrObjectName.ObjectName));
    //    //if (instrObjectNameVar == null)
    //    //{
    //    //    result.AddError(ErrorCode.ParserVarNotDefined, instrObjectName.FirstScriptToken());
    //    //    return false;
    //    //}

    //    // check the final value of the var, can be a value, a fct call or a math expr
    //    InstrSetVar instrSetVar = program.FindLastVarSet(instrObjectName.ObjectName);
    //    if (instrSetVar == null)
    //    {
    //        result.AddError(ErrorCode.ParserVarNotDefined, instrObjectName.FirstScriptToken());
    //        return false;
    //    }

    //    // its a final int value, check it
    //    if (instrSetVar.InstrRight.InstrType== InstrType.Value)
    //    {
    //        var instrValue = instrSetVar.InstrRight as InstrValue;
    //        if (!InstrUtils.GetValueIntFromInstrValue(result, instrValue, scriptLineNum, out int val))
    //            return false;

    //        // check the int value, should be >= 1
    //        if (val < 1)
    //        {
    //            result.AddError(ErrorCode.ParserValueIntWrong, instrValue.FirstScriptToken());
    //            return false;
    //        }
    //        // save the var object
    //        //instrOnExcel.CurrOnSheet.InstrFirstDataRow = instrObjectName;
    //        return true;
    //    }

    //    // can be fct call, a math expr
    //    // TODO: 

    //    result.AddError(ErrorCode.ParserCaseNotManaged, instrSetVar.InstrRight.FirstScriptToken());
    //    return false;

    //}
}