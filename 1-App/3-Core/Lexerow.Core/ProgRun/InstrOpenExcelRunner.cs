using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;
using Lexerow.Core.System.Excel;
using Lexerow.Core.System.ProgRun;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.ProgRun;

/// <summary>
/// Instr OpenExcel runner.
/// </summary>
public class InstrOpenExcelRunner
{
    IActivityLogger _logger;

    IExcelProcessor _excelProcessor;

    public InstrOpenExcelRunner(IActivityLogger activityLogger, IExcelProcessor excelProcessor)
    {
        _logger = activityLogger;
        _excelProcessor = excelProcessor;
    }

    /// <summary>
    /// Execute instr OpenExcel.
    /// Format:
    ///   OpenExcel("file.xlsx")
    ///   OpenExcel(fileName)
    /// </summary>
    /// <param name="execResult"></param>
    /// <param name="ctx"></param>
    /// <param name="listInstr"></param>
    /// <param name="listVar"></param>
    /// <param name="instr"></param>
    /// <returns></returns>
    public bool Run(ExecResult execResult, ProgramRunnerContext ctx, List<ExecVar> listVar, InstrOpenExcel instr)
    {
        _logger.LogRunStart(ActivityLogLevel.Info, "ProgRunner.ExecInstrOpenExcel", string.Empty);

        // param is a ObjectName ? exp: OpenExcel(fileName)
        InstrObjectName instrObjectName = instr.Param as InstrObjectName;
        if (instrObjectName != null)
        {
            // get the var name, should be defined before
            ExecVar execVar = listVar.FirstOrDefault(v => v.NameEquals(instrObjectName.ObjectName));
            if (execVar == null)
            {
                execResult.AddError(ErrorCode.RunInstrVarNotFound, instr.Param.FirstScriptToken());
                return false;
            }

            // execute OpenExcel
            return Run(execResult, ctx, instr, execVar.GetValueString());
        }

        // param is constValue type string? exp: OpenExcel("file.xlsx")
        InstrConstValue instrConstValue = instr.Param as InstrConstValue;
        if (instrConstValue != null)
        {
            // execute the instr OpenExcel(fileName)
            return Run(execResult, ctx, instr, instrConstValue.RawValue);
        }

        // need to execute the instr param (fct call? ...)
        ctx.StackInstr.Push(instr.Param);
        return true;
    }

    public bool Run(ExecResult execResult, ProgramRunnerContext ctx, InstrBase instr, string fileName)
    {
        // execute the instr OpenExcel(fileName)
        if (!_excelProcessor.Open(fileName, out IExcelFile excelFile, out ExecResultError error))
        {
            execResult.AddError(error);
            return false;
        }

        InstrExcelFileObject instrExcelFileObject = new InstrExcelFileObject(instr.FirstScriptToken(), fileName, excelFile);
        ctx.PrevInstrExecuted = instrExcelFileObject;
        // now remove the OpenExcel from instr to execute, it's done
        ctx.StackInstr.Pop();
        return true;
    }

}
