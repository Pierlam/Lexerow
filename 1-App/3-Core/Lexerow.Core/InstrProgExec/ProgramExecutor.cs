using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;
using Lexerow.Core.System.InstrDef;
using Lexerow.Core.System.InstrDef.Object;
using OpenExcelSdk;
using System.Diagnostics;

namespace Lexerow.Core.InstrProgExec;

/// <summary>
/// Program script executor.
/// </summary>
public class ProgramExecutor
{
    private IActivityLogger _logger;

    private ExcelProcessor _excelProcessor;

    private ProgExecVarMgr _progExecVarMgr = new ProgExecVarMgr();

    private InstrExecutor _instrExecutor;

    public ProgramExecutor(IActivityLogger activityLogger, ExcelProcessor excelProcessor)
    {
        _logger = activityLogger;
        _excelProcessor = excelProcessor;

        _instrExecutor = new InstrExecutor(activityLogger, _excelProcessor, _progExecVarMgr);
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

        _logger.LogExec(ActivityLogLevel.Info, "ProgramExecutor.Exec.Start", "Name: " + program.Script.Name + ", Instr Count: " + program.ListInstr.Count.ToString());
        Stopwatch stopwatch = new Stopwatch();
        stopwatch.Start();

        // execute instr, one by one
        foreach (var instrBase in program.ListInstr)
        {
            res = _instrExecutor.ExecInstr(result, program, _progExecVarMgr, instrBase);
            if (!res) return false;
        }

        // close opened excel file
        res&=CloseOpenedExcelFiles(result, _progExecVarMgr);

        stopwatch.Stop();
        string elapsedTime = string.Format("{0:hh\\:mm\\:ss}", stopwatch.Elapsed);
        if (!res)
            _logger.LogExecError(result.ListError[0], "ProgramExecutor.Exec", "Error count: " + result.ListError.Count.ToString());
        else
            _logger.LogExec(ActivityLogLevel.Info, "ProgramExecutor.Exec", "Elapsed time: " + elapsedTime);
        return res;
    }

    /// <summary>
    /// Close opened excel file
    /// e.g. file=CreateExcel()
    /// </summary>
    /// <param name="result"></param>
    /// <param name="progExecVarMgr"></param>
    /// <returns></returns>
    private bool CloseOpenedExcelFiles(Result result, ProgExecVarMgr progExecVarMgr)
    {
        bool res = true;
        foreach(ProgExecVar execVar in progExecVarMgr.ListExecVar)
        {
            InstrObjectExcelFile instrObjectExcelFile = execVar.Value as InstrObjectExcelFile;
            if (instrObjectExcelFile!=null && instrObjectExcelFile.ExcelFile!=null)
            {
                res &= CloseFileExecutor.Exec(result, _excelProcessor, instrObjectExcelFile.ExcelFile);
            }
        }

        return res;
    }
}