using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;
using Lexerow.Core.System.Excel;
using System.Diagnostics;

namespace Lexerow.Core.ProgExec;

/// <summary>
/// Program script executor.
/// </summary>
public class ProgramExecutor
{
    IActivityLogger _logger;

    IExcelProcessor _excelProcessor;

    ProgExecVarMgr _progRunVarMgr=new ProgExecVarMgr();
    InstrExecutor _instrExecutor;

    public ProgramExecutor(IActivityLogger activityLogger, IExcelProcessor excelProcessor)
    {
        _logger = activityLogger;
        _excelProcessor = excelProcessor;

        _instrExecutor= new InstrExecutor(activityLogger, _excelProcessor);
    }

    /// <summary>
    /// Execute a program, obtained after the compilation of a script.
    /// </summary>
    /// <param name="execResult"></param>
    /// <param name="programScript"></param>
    /// <returns></returns>
    public bool Exec(ExecResult execResult, ProgramScript programScript)
    {
        bool res = true;

        _logger.LogExecStart(ActivityLogLevel.Important, "ProgramExecutor.Exec", "Name: " + programScript.Script.Name + ", Instr Count: " + programScript.ListInstr.Count.ToString());
        Stopwatch stopwatch = new Stopwatch();
        stopwatch.Start();

        // convert program instr to instr to execute

        // execute instr, one by one        
        foreach (var instrBase in programScript.ListInstr)
        {
            res = _instrExecutor.ExecInstr(execResult, _progRunVarMgr, instrBase);
            if (!res) return false;
        }

        // close all opened excel file, if its not done
        // TODO: don't call it! each excel file will be closed just the use of it at the end of the exec of the instr OnExcel
        //res= CloseAllOpenedExcelFile(execResult, _listExecVar);

        stopwatch.Stop();
        string elapsedTime = string.Format("{0:hh\\:mm\\:ss}", stopwatch.Elapsed);
        if (!res)
            _logger.LogExecEndError(execResult.ListError[0], "ProgramExecutor.Exec", "Error count: " + execResult.ListError.Count.ToString());
        else
            _logger.LogExecEnd(ActivityLogLevel.Important, "ProgramExecutor.Exec", "Elapsed time: " + elapsedTime);
        return res;
    }
}
