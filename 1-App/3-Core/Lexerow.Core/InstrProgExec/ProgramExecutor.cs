using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;
using Lexerow.Core.System.Excel;
using Lexerow.Core.System.InstrDef;
using System.Diagnostics;

namespace Lexerow.Core.InstrProgExec;

/// <summary>
/// Program script executor.
/// </summary>
public class ProgramExecutor
{
    private IActivityLogger _logger;

    private IExcelProcessor _excelProcessor;

    private ProgExecVarMgr _progExecVarMgr = new ProgExecVarMgr();

    private InstrExecutor _instrExecutor;

    public ProgramExecutor(IActivityLogger activityLogger, IExcelProcessor excelProcessor)
    {
        _logger = activityLogger;
        _excelProcessor = excelProcessor;

        _instrExecutor = new InstrExecutor(activityLogger, _excelProcessor);
    }

    /// <summary>
    /// Execute a program, obtained after the compilation of a script.
    /// </summary>
    /// <param name="result"></param>
    /// <param name="program"></param>
    /// <returns></returns>
    public bool Exec(Result result, Program program)
    {
        bool res = true;

        _logger.LogExecStart(ActivityLogLevel.Important, "ProgramExecutor.Exec", "Name: " + program.Script.Name + ", Instr Count: " + program.ListInstr.Count.ToString());
        Stopwatch stopwatch = new Stopwatch();
        stopwatch.Start();

        // convert program instr to instr to execute

        // execute instr, one by one
        foreach (var instrBase in program.ListInstr)
        {
            res = _instrExecutor.ExecInstr(result, program, _progExecVarMgr, instrBase);
            if (!res) return false;
        }

        // close all opened excel file, if its not done
        // TODO: don't call it! each excel file will be closed just the use of it at the end of the exec of the instr OnExcel
        //res= CloseAllOpenedExcelFile(result, _listExecVar);

        stopwatch.Stop();
        string elapsedTime = string.Format("{0:hh\\:mm\\:ss}", stopwatch.Elapsed);
        if (!res)
            _logger.LogExecEndError(result.ListError[0], "ProgramExecutor.Exec", "Error count: " + result.ListError.Count.ToString());
        else
            _logger.LogExecEnd(ActivityLogLevel.Important, "ProgramExecutor.Exec", "Elapsed time: " + elapsedTime);
        return res;
    }
}