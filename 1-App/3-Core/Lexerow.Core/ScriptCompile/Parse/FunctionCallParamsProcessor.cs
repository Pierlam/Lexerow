using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;
using Lexerow.Core.System.InstrDef;
using Lexerow.Core.System.ScriptCompile;
using Lexerow.Core.System.ScriptDef;

namespace Lexerow.Core.ScriptCompile.Parse;

internal class FunctionCallParamsProcessor
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
    public static bool ProcessFunctionCallParams(IActivityLogger logger, Result result, List<InstrObjectName> listVar, CompilStackInstr stackInstr, ScriptToken scriptToken, List<InstrBase> listInstrToExec, List<InstrBase> listParams)
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

        logger.LogCompilStart(ActivityLogLevel.Important, "FunctionCallParamsProcessor.ProcessFunctionCallParams", "InstrType: " + instrBase.InstrType);

        if (instrBase.InstrType == InstrType.FuncSelectFiles)
            return ProcessSelectFiles(logger, result, listVar, instrBase as InstrFuncSelectFiles, listInstrToExec, listParams);

        // get the last instr from the stack

        throw new NotImplementedException("not yet implemented, InstrType:" + instrBase.InstrType.ToString());
    }

    private static bool ProcessSelectFiles(IActivityLogger logger, Result result, List<InstrObjectName> listVar, InstrFuncSelectFiles instr, List<InstrBase> listInstrToExec, List<InstrBase> listParams)
    {
        logger.LogCompilStart(ActivityLogLevel.Info, "FunctionCallParamsProcessor.ProcessSelectFiles", "Param count: " + instr.ListInstrParams.Count);
        // only one param expected, type should be string or an instr returning a string
        if (listParams.Count != 1)
        {
            result.AddError(ErrorCode.ParserFctParamCountWrong, instr.ListScriptToken[0], listParams.Count.ToString());
            return false;
        }

        //--is the param a string const value token?  exp: SelectFiles("MyFile.xlsx")
        InstrValue instrValue = listParams[0] as InstrValue;
        if (instrValue != null)
        {
            // the const value type should be a string
            if (instrValue.ValueBase.ValueType == System.ValueType.String)
            {
                // push the string param to the instr SelectFiles
                instr.AddParamSelect(instrValue);
                return true;
            }

            // not a string, error
            result.AddError(ErrorCode.ParserFctParamTypeWrong, instr.ListScriptToken[0], listParams.Count.ToString());
            return false;
        }

        //--is the param a varName source code token?  exp: SelectFiles(fileName)
        InstrObjectName instrObjectName = listParams[0] as InstrObjectName;
        if (instrObjectName != null)
        {
            // check that the var is defined
            if (listVar.FirstOrDefault(x => x.ObjectName.Equals(instrObjectName.ObjectName, StringComparison.InvariantCultureIgnoreCase)) == null)
            {
                // not a string, error
                result.AddError(ErrorCode.ParserFctParamVarNotDefined, instr.ListScriptToken[0], listParams.Count.ToString());
                return false;
            }

            // push the string param to the instr OpenExcel
            instr.AddParamSelect(instrObjectName);
            return true;
        }

        result.AddError(ErrorCode.ParserFctParamTypeWrong, instr.ListScriptToken[0], listParams[0].GetType().ToString());
        return false;
    }
}