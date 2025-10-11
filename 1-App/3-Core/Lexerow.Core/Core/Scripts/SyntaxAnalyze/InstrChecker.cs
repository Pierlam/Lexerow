using Lexerow.Core.System;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.Formula.Functions;
using Org.BouncyCastle.Utilities.Collections;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Scripts.SyntaxAnalyze;
internal class InstrChecker
{
    /// <summary>
    /// Check the fct call: params set and type ok?

    /// </summary>
    /// <param name="execResult"></param>
    /// <param name="instrBase"></param>
    /// <returns></returns>
    public static bool CheckFunctionCall(ExecResult execResult, InstrBase instrBase)
    {
        // not a fct call, bye
        if (!instrBase.IsFunctionCall) return true;

        //--is it OpenExcel?
        InstrOpenExcel instrOpenExcel = instrBase as InstrOpenExcel;
        if (instrOpenExcel != null) 
        {
            if (instrOpenExcel.Param == null)
            {
                execResult.AddError(ErrorCode.SyntaxAnalyzerFctParamCountWrong, instrBase.FirstScriptToken());
                return false;
            }
            // isntr OpenExcel is ok
            return true;
        }

        //--fct not managed 
        throw new NotImplementedException(instrBase.ToString());
    }

}
