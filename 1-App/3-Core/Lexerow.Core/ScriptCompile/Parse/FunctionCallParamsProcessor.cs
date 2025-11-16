using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;
using Lexerow.Core.System.ScriptCompile;
using Lexerow.Core.System.ScriptDef;

namespace Lexerow.Core.ScriptCompile.Parse;

internal class FunctionCallParamsProcessor
{
    /// <summary>
    /// Manage function call parameters.
    /// check that parameters match the function call.
    /// </summary>
    /// <param name="execResult"></param>
    /// <param name="stackInstr"></param>
    /// <param name="scriptToken"></param>
    /// <param name="listInstrToExec"></param>
    /// <param name="listParams"></param>
    /// <returns></returns>
    public static bool ProcessFunctionCallParams(IActivityLogger logger, ExecResult execResult, List<InstrObjectName> listVar, CompilStackInstr stackInstr, ScriptToken scriptToken, List<InstrBase> listInstrToExec, List<InstrBase> listParams)
    {
        // the stack is empty?
        if (stackInstr.Count == 0)
        {
            // function call name expected
            execResult.AddError(ErrorCode.ParserFctNameExpected, scriptToken);

            return false;
        }

        // read the last instr from the stack
        InstrBase instrBase = stackInstr.Peek();

        logger.LogCompilStart(ActivityLogLevel.Important, "FunctionCallParamsProcessor.ProcessFunctionCallParams", "InstrType: " + instrBase.InstrType);

        if (instrBase.InstrType == InstrType.SelectFiles)
            return ProcessSelectFiles(logger, execResult, listVar, instrBase as InstrSelectFiles, listInstrToExec, listParams);

        // get the last instr from the stack

        throw new NotImplementedException("not yet implemented, InstrType:" + instrBase.InstrType.ToString());
    }

    private static bool ProcessSelectFiles(IActivityLogger logger, ExecResult execResult, List<InstrObjectName> listVar, InstrSelectFiles instr, List<InstrBase> listInstrToExec, List<InstrBase> listParams)
    {
        logger.LogCompilStart(ActivityLogLevel.Info, "FunctionCallParamsProcessor.ProcessSelectFiles", "Param count: " + instr.ListInstrParams.Count);
        // only one param expected, type should be string or an instr returning a string
        if (listParams.Count != 1)
        {
            execResult.AddError(ErrorCode.ParserFctParamCountWrong, instr.ListScriptToken[0], listParams.Count.ToString());
            return false;
        }

        //--is the param a string const value token?  exp: SelectFiles("MyFile.xlsx")
        InstrConstValue instrConstValue = listParams[0] as InstrConstValue;
        if (instrConstValue != null)
        {
            // the const value type should be a string
            if (instrConstValue.ValueBase.ValueType == System.ValueType.String)
            {
                // push the string param to the instr SelectFiles
                instr.AddParamSelect(instrConstValue);
                return true;
            }

            // not a string, error
            execResult.AddError(ErrorCode.ParserFctParamTypeWrong, instr.ListScriptToken[0], listParams.Count.ToString());
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
                execResult.AddError(ErrorCode.ParserFctParamVarNotDefined, instr.ListScriptToken[0], listParams.Count.ToString());
                return false;
            }

            // push the string param to the instr OpenExcel
            instr.AddParamSelect(instrObjectName);
            return true;
        }

        execResult.AddError(ErrorCode.ParserFctParamTypeWrong, instr.ListScriptToken[0], listParams[0].GetType().ToString());
        return false;
    }
}