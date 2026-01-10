using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;
using Lexerow.Core.System.InstrDef;
using Lexerow.Core.System.InstrDef.FuncCall;
using Lexerow.Core.System.ScriptCompile;
using Lexerow.Core.System.ScriptDef;
using Lexerow.Core.Utils;

namespace Lexerow.Core.ScriptCompile.Parse;

internal class FunctionCallParamsParser
{
    /// <summary>
    /// Manage function call parameters.
    /// check that parameters match the function call.
    /// </summary>
    /// <param name="result"></param>
    /// <param name="stackInstr"></param>
    /// <param name="scriptToken"></param>
    /// <param name="listInstrToExec"></param>
    /// <param name="listParams"></param>
    /// <returns></returns>
    public static bool ProcessFunctionCallParams(IActivityLogger logger, Result result, List<InstrNameObject> listVar, CompilStackInstr stackInstr, ScriptToken scriptToken, Program program, List<InstrBase> listParams)
    {
        // the stack is empty?
        if (stackInstr.Count == 0)
        {
            // function call name expected
            result.AddError(ErrorCode.ParserFctNameExpected, scriptToken);
            return false;
        }

        // read the last instr from the stack
        InstrBase instrBase = stackInstr.Peek();

        logger.LogCompilStart(ActivityLogLevel.Info, "FunctionCallParamsParser.ProcessFunctionCallParams", "InstrType: " + instrBase.InstrType);

        //-item before openBracket is an object name? exp: fct()
        if (!instrBase.IsFunctionCall)
        {
            result.AddError(ErrorCode.ParserTokenNotExpected, instrBase.FirstScriptToken());
            return false;
        }

        if (instrBase.InstrType == InstrType.FuncSelectFiles)
            return ProcessFuncSelectFiles(logger, result, listVar, instrBase as InstrFuncCallSelectFiles, program, listParams);

        if (instrBase.InstrType == InstrType.FuncDate)
            return ProcessFuncDate(logger, result, listVar, instrBase as InstrFuncCallDate, program, listParams);

        // function call name not expected
        result.AddError(ErrorCode.ParserFctNameNotExpected, scriptToken);
        return false;
    }

    private static bool ProcessFuncSelectFiles(IActivityLogger logger, Result result, List<InstrNameObject> listVar, InstrFuncCallSelectFiles instr, Program program, List<InstrBase> listParams)
    {
        logger.LogCompilStart(ActivityLogLevel.Debug, "FunctionCallParamsParser.ProcessFuncSelectFiles", "Param count IN: " + listParams.Count);

        // only one param expected, type should be string or an instr returning a string
        if (listParams.Count != 1)
        {
            result.AddError(ErrorCode.ParserFctParamCountWrong, instr.ListScriptToken[0], listParams.Count.ToString());
            return false;
        }

        //--exp: SelectFiles("MyFile.xlsx") or SelectFiles(filename) or SelectFiles(fct())

        if (!InstrUtils.GetStringFromInstrParser(result, program, listParams[0], out _, out _))
            return false;

        instr.AddParamSelect(listParams[0]);
        return true;
    }

    private static bool ProcessFuncDate(IActivityLogger logger, Result result, List<InstrNameObject> listVar, InstrFuncCallDate instr, Program program, List<InstrBase> listParams)
    {
        logger.LogCompilStart(ActivityLogLevel.Debug, "FunctionCallParamsParser.ProcessFuncSelectFiles", "Param count In: " + listParams.Count);

        // 3 param expected, type should be an int or an instr returning an int
        if (listParams.Count != 3)
        {
            result.AddError(ErrorCode.ParserFctParamCountWrong, instr.ListScriptToken[0], listParams.Count.ToString());
            return false;
        }

        // process 1st param: year
        if (!InstrUtils.GetIntFromInstrParser(result, program, listParams[0], out bool yearSet, out int year))
            return false;
        instr.InstrYear = listParams[0];

        // process 2nd param: month
        if (!InstrUtils.GetIntFromInstrParser(result, program, listParams[1], out bool monthSet, out int month))
            return false;
        instr.InstrMonth = listParams[1];

        // process 3rd param: day
        if (!InstrUtils.GetIntFromInstrParser(result, program, listParams[2], out bool daySet, out int day))
            return false;
        instr.InstrDay = listParams[2];

        if (!yearSet || !monthSet || !daySet)
            // not able to check the date here, need to wait the execution
            return true;

        // 3 values are found, check the date
        try
        {
            var date = new DateOnly(year, month, day);
            return true;
        }
        catch (Exception ex)
        {
            result.AddError(ErrorCode.ParserFctParamWrong, instr.ListScriptToken[0], "year: " + year + ", month: " + month + ", day:" + day);
            return false;
        }
    }
}