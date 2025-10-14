using Lexerow.Core.System;
using Lexerow.Core.System.Excel;
using Lexerow.Core.System.Exec.Event;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Core.Exec;

/// <summary>
/// Execute functions manager.
/// </summary>
public class ExecFunctionMgr
{
    IExcelProcessor _excelProcessor;
    public ExecFunctionMgr(IExcelProcessor excelProcessor)
    {
        _excelProcessor= excelProcessor;
    }

    public Action<AppTrace> AppTraceEvent { get; set; }

    public bool ExecFunction(ExecResult execResult, Stack<InstrBase> stackInstr, InstrBase instrBaseFunction, List<InstrBase> listFuncParams, List<ExecVar> listExecVar, DateTime execStart)
    {
        //--is it OpenExcel?
        if (instrBaseFunction.InstrType == InstrType.OpenExcel)
        {
            // should have one parameter, a filename
            if (listFuncParams.Count != 1)
            {
                execResult.AddError(new ExecResultError(ErrorCode.FuncOneParamExpected, InstrType.OpenExcel.ToString(), listFuncParams.Count.ToString()));
                return false;
            }
            // check
            InstrConstValue instrConstValue = listFuncParams[0] as InstrConstValue;
            // TODO: if =null

            if (!_excelProcessor.Open(instrConstValue.RawValue, out IExcelFile excelFile, out ExecResultError error))
            {
                execResult.AddError(error);
                SendAppTraceExec(AppTraceLevel.Error, "ExecInstrOpenExcelFile", InstrOpenExcelExecEvent.CreateFinished(execStart, InstrBaseExecEventResult.Error, instrConstValue.RawValue));
                return false;
            }

            // create instruction : ExcelFileObject and save it in the stack
            InstrExcelFileObject excelFileObject = new InstrExcelFileObject(instrConstValue.FirstScriptToken(), instrConstValue.RawValue, excelFile);
            stackInstr.Push(excelFileObject);
            return true;
        }

        //--function not implemented or unknow
        execResult.AddError(new ExecResultError(ErrorCode.FuncNotExists, instrBaseFunction.InstrType.ToString()));
        return false;
    }

    public void SendAppTraceExec(AppTraceLevel level, string msg, InstrBaseExecEvent execEvent)
    {
        if (AppTraceEvent == null) return;

        AppTrace appTrace = new AppTrace(AppTraceModule.Exec, level, msg, execEvent);
        AppTraceEvent(appTrace);
    }

}
