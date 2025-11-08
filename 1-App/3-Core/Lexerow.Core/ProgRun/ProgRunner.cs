using Lexerow.Core.Core.Exec;
using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;
using Lexerow.Core.System.Excel;
using Lexerow.Core.System.ProgRun;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.Formula.Functions;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.ProgRun;
public class ProgramRunner
{
    IActivityLogger _logger;

    IExcelProcessor _excelProcessor;

    //List<ProgRunVar> _listExecVar = new List<ProgRunVar>();
    ProgRunVarMgr _progRunVarMgr=new ProgRunVarMgr();
    InstrRunner _instrRunner;



    public ProgramRunner(IActivityLogger activityLogger, IExcelProcessor excelProcessor)
    {
        _logger = activityLogger;
        _excelProcessor = excelProcessor;

        _instrRunner= new InstrRunner(activityLogger, _excelProcessor);
    }

    /// <summary>
    /// Execute a program, obtained after the compilation of a script.
    /// </summary>
    /// <param name="execResult"></param>
    /// <param name="programScript"></param>
    /// <returns></returns>
    public bool Run(ExecResult execResult, ProgramScript programScript)
    {
        bool res = true;

        _logger.LogRunStart(ActivityLogLevel.Important, "ProgRunner.Run", "Name: " + programScript.Script.Name + ", Instr Count: " + programScript.ListInstr.Count.ToString());
        Stopwatch stopwatch = new Stopwatch();
        stopwatch.Start();

        // execute instr, one by one        
        foreach (var instrBase in programScript.ListInstr)
        {
            res = _instrRunner.ExecInstr(execResult, _progRunVarMgr, instrBase);
            if (!res) return false;
        }

        // close all opened excel file, if its not done
        // TODO: don't call it! each excel file will be closed just the use of it at the end of the exec of the instr OnExcel
        //res= CloseAllOpenedExcelFile(execResult, _listExecVar);

        stopwatch.Stop();
        string elapsedTime = string.Format("{0:hh\\:mm\\:ss}", stopwatch.Elapsed);
        if (!res)
            _logger.LogRunEndError(execResult.ListError[0], "ProgRunner.Run", "Error count: " + execResult.ListError.Count.ToString());
        else
            _logger.LogRunEnd(ActivityLogLevel.Important, "ProgRunner.Run", "Elapsed time: " + elapsedTime);
        return res;
    }


    /// <summary>
    /// Close all opened excel file, if its not done.
    /// </summary>
    /// <param name="listExecVar"></param>
    bool CloseAllOpenedExcelFile(ExecResult execResult, List<ProgRunVar> listExecVar)
    {
        _logger.LogRunStart(ActivityLogLevel.Info, "ProgRunner.CloseAllOpenedExcelFile", string.Empty);
        bool res = true;
        foreach (ProgRunVar execVar in listExecVar)
        {
            InstrExcelFileObject instrExcelFileObject = execVar.Value as InstrExcelFileObject;
            if (instrExcelFileObject != null)
            {
                if (!CloseExcelFileRunner.Exec(_excelProcessor, instrExcelFileObject.ExcelFile, out ExecResultError error))
                {
                    execResult.AddError(error);
                    res = false;
                }
                instrExcelFileObject.ExcelFile = null;
            }
        }

        return res;
    }

}
