using Lexerow.Core.System;
using Lexerow.Core.System.InstrDef;
using Lexerow.Core.System.InstrDef.Func;

namespace Lexerow.Core.ScriptCompile.Parse;

internal class InstrChecker
{
    /// <summary>
    /// Check the fct call: params set and type ok?
    /// </summary>
    /// <param name="result"></param>
    /// <param name="instrBase"></param>
    /// <returns></returns>
    public static bool CheckFunctionCall(Result result, InstrBase instrBase)
    {
        // not a fct call, bye
        if (!instrBase.IsFunctionCall) return true;

        //--is it SelectFiles?
        InstrFuncSelectFiles instrFuncSelectFiles = instrBase as InstrFuncSelectFiles;
        if (instrFuncSelectFiles != null)
        {
            // need at least one parameter
            if (instrFuncSelectFiles.ListInstrParams.Count == 0)
            {
                result.AddError(ErrorCode.ParserFctParamCountWrong, instrBase.FirstScriptToken());
                return false;
            }
            return true;
        }

        //--is it Date?
        InstrFuncDate instrFuncDate = instrBase as InstrFuncDate;
        if(instrFuncDate!=null)
        {
            if(instrFuncDate.InstrYear==null || instrFuncDate.InstrMonth==null || instrFuncDate.InstrDay==null)
            {
                result.AddError(ErrorCode.ParserFctParamWrong, instrBase.FirstScriptToken());
                return false;
            }
            return true;
        }
        //--fct not managed
        result.AddError(ErrorCode.ParserFctNotManaged, instrBase.FirstScriptToken());
        return false;
    }
}