using Lexerow.Core.System;
using Lexerow.Core.System.Compilator;
using NPOI.SS.Formula.Functions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Scripts.SyntaxAnalyze;
internal class FunctionCallParamsProcessor
{
    /// <summary>
    /// Manage function call parameters.
    /// check that parameters match the function call.
    /// </summary>
    /// <param name="execResult"></param>
    /// <param name="stkItems"></param>
    /// <param name="scriptToken"></param>
    /// <param name="listInstrToExec"></param>
    /// <param name="listParams"></param>
    /// <returns></returns>
    public static bool ProcessFunctionCallParams(ExecResult execResult, List<InstrObjectName> listVar, Stack<InstrBase> stkItems, ScriptToken scriptToken, List<InstrBase> listInstrToExec, List<InstrBase> listParams)
    {
        // the stack is empty? 
        if (stkItems.Count == 0)
        {
            // function call name expected
            execResult.AddError(ErrorCode.SyntaxAnalyzerFunctionCallNameExpected, scriptToken);

            return false;
        }

        // read the last instr from the stack
        InstrBase instrBase = stkItems.Peek();

        if (instrBase.InstrType == InstrType.OpenExcel)
            return ProcessOpenExcel(execResult, listVar, instrBase as InstrOpenExcel, listInstrToExec, listParams);

        // get the last instr from the stack

        throw new NotImplementedException("not yet implemented, InstrType:" + instrBase.InstrType.ToString());

    }

    static bool ProcessOpenExcel(ExecResult execResult, List<InstrObjectName> listVar, InstrOpenExcel instr, List<InstrBase> listInstrToExec, List<InstrBase> listParams)
    {
        // only one param expected, type should be string or an instr returning a string
        if (listParams.Count != 1)
        {
            execResult.AddError(ErrorCode.SyntaxAnalyzerFctParamCountWrong, instr.ListScriptToken[0], listParams.Count.ToString());
            return false;
        }

        //--is the param a string const value token?  exp: OpenExcel("MyFile.xlsx")
        InstrConstValue instrConstValue = listParams[0] as InstrConstValue;
        if(instrConstValue!=null)
        {
            // the const value type should be a string 
            if (instrConstValue.ValueBase.ValueType == System.ValueType.String)
            {
                // push the string param to the instr OpenExcel
                instr.Param = instrConstValue;
                return true;
            }

            // not a string, error
            execResult.AddError(ErrorCode.SyntaxAnalyzerFctParamTypeWrong, instr.ListScriptToken[0], listParams.Count.ToString());
            return false;
        }

        //--is the param a varName source code token?  exp: OpenExcel(fileName)
        InstrObjectName instrObjectName = listParams[0] as InstrObjectName;
        if (instrObjectName != null) 
        {
            // check that the var is defined
            if(listVar.FirstOrDefault(x => x.ObjectName.Equals(instrObjectName.ObjectName, StringComparison.InvariantCultureIgnoreCase))==null)
            {
                // not a string, error
                execResult.AddError(ErrorCode.SyntaxAnalyzerFctParamVarNotDefined, instr.ListScriptToken[0], listParams.Count.ToString());
                return false;
            }

            // push the string param to the instr OpenExcel
            instr.Param = instrObjectName;
            return true;
        }


        execResult.AddError(ErrorCode.SyntaxAnalyzerFctParamTypeWrong, instr.ListScriptToken[0], listParams[0].GetType().ToString());
        return false;
    }
}
