using Lexerow.Core.System;

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

        //--is it OpenExcel?
        InstrSelectFiles instrOpenExcel = instrBase as InstrSelectFiles;
        if (instrOpenExcel != null)
        {
            // need at least one parameter
            if (instrOpenExcel.ListInstrParams.Count == 0)
            {
                result.AddError(ErrorCode.ParserFctParamCountWrong, instrBase.FirstScriptToken());
                return false;
            }
            // isntr OpenExcel is ok
            return true;
        }

        //--fct not managed
        throw new NotImplementedException(instrBase.ToString());
    }
}