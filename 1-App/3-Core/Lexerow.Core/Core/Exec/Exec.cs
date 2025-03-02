using Lexerow.Core.System;
using Lexerow.Core.System.Excel;
using Lexerow.Core.System.Exec.Event;
using Lexerow.Core.Utils;
using Microsoft.Extensions.Logging;
using NPOI.HPSF;
using NPOI.SS.Extractor;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core;

/// <summary>
/// Execute instruction
/// </summary>
public class Exec
{
    ILoggerFactory _loggerFactory;

    IExcelProcessor _excelProcessor;

    CoreData _coreData;

    List<ExecVar> _listExecVar = new List<ExecVar>();

    DateTime _execStart;

    DateTime _execStartCurrInstr;

    /// <summary>
    /// Constructor.
    /// </summary>
    /// <param name="loggerFactory"></param>
    /// <param name="coreData"></param>
    /// <param name="excelProcessor"></param>
    public Exec(ILoggerFactory loggerFactory, CoreData coreData, IExcelProcessor excelProcessor)
    {
        _loggerFactory = loggerFactory;
        _coreData = coreData;
        _excelProcessor = excelProcessor;
    }

    /// <summary>
    /// When a rule is fired.
    /// </summary>
    public Action<InstrBaseExecEvent> EventOccurs { get; set; }

    public bool AddInstr(InstrBase instrBase)
    {
        _coreData.ListInstr.Add(instrBase);
        return true;
    }

    public ExecResult Compile()
    {
        ExecResult execResult = new ExecResult();

        // TODO: check all instruction, one by one 
        return execResult;
    }

    /// <summary>
    /// Execute all saved instruction.
    /// Need to compile instr before execute them.
    /// </summary>
    /// <returns></returns>
    public ExecResult Execute()
    {
        _execStart= DateTime.Now;   

        ExecResult execResult;

        foreach (var instr in _coreData.ListInstr) 
        {
            if (instr.InstrType == InstrType.OpenExcel)
            {
                _execStartCurrInstr= DateTime.Now;
                execResult= ExecInstrOpenExcelFile(instr as InstrOpenExcel, _listExecVar, _execStartCurrInstr);
                if(!execResult.Result)
                    return execResult;

                continue;
            }

            if(instr.InstrType == InstrType.CloseExcel)
            {
                // TODO: ExecInstrCloseExcelFile
                continue;
            }

            if(instr.InstrType == InstrType.ForEachRowIfThen)
            {
                _execStartCurrInstr = DateTime.Now;
                execResult = ExecInstrForEachRowIfThen(instr as InstrForEachRowIfThen, _listExecVar, _execStartCurrInstr);
                if (!execResult.Result)
                    return execResult;
 
                continue;
            }
        }

        // close all opened excel file, if its not done
        CloseAllOpenedExcelFile(_listExecVar);

        return new ExecResult(); 
    }

    /// <summary>
    /// Execute the instr OpenExcel.
    /// return an object: an excel file.
    /// </summary>
    /// <param name="instrOpenExcel"></param>
    /// <returns></returns>
    ExecResult ExecInstrOpenExcelFile(InstrOpenExcel instrOpenExcel, List<ExecVar> listExecVar, DateTime execStart)
    {
        ExecResult result = new ExecResult();
        FireEvent(InstrOpenExcelExecEvent.CreateStart(instrOpenExcel.FileName));

        if (!_excelProcessor.Open(instrOpenExcel.FileName, out IExcelFile excelFile, out CoreError error))
        {
            result.AddError(error);
            FireEvent(InstrOpenExcelExecEvent.CreateFinished(execStart, InstrBaseExecEventResult.Error, instrOpenExcel.FileName));
            return result;
        }

        ExecVar execVar = new ExecVar(instrOpenExcel.ExcelFileObjectName, ExecVarType.ExcelFile, excelFile);
        listExecVar.Add(execVar);

        FireEvent(InstrOpenExcelExecEvent.CreateFinished(execStart, InstrBaseExecEventResult.Ok, instrOpenExcel.FileName));
        return result;
    }

    /// <summary>
    /// Execute instr: 
    /// 	ForEachRowIfThen(file, sheetNum, A,Value>10, Value=10)
    /// ForEach Cell c In Cols(A)
    ///    If c.Value >10 Then c.Value=10
    /// </summary>
    /// <param name="instr"></param>
    /// <param name="listExecVar"></param>
    /// <returns></returns>
    ExecResult ExecInstrForEachRowIfThen(InstrForEachRowIfThen instr, List<ExecVar> listExecVar, DateTime execStart)
    {
        ExecResult execResult = new ExecResult();
        FireEvent(InstrForEachRowIfThenExecEvent.CreateStart());

        // get the file object by name
        ExecVar execVar = listExecVar.FirstOrDefault(ev => ev.Name.Equals(instr.ExcelFileObjectName, StringComparison.InvariantCultureIgnoreCase));
        if(execVar==null)
        {
            FireEvent(InstrForEachRowIfThenExecEvent.CreateFinished(execStart,InstrBaseExecEventResult.Error));
            execResult.AddError(new CoreError(ErrorCode.ExcelFileObjectNameDoesNotExists, null));
            return execResult;
        }

        // check that the var value in an excel file object
        if (execVar.ExecVarType!= ExecVarType.ExcelFile)
        {
            // TODO: erreur the type of the var is wrong            
        }

        // execute the instr on each cols, on each datarow
        execResult = ExecInstrForEachRowIfThenMgr.Exec(FireEvent, execStart, _excelProcessor,  execVar.Value as IExcelFile, instr);

        return execResult;
    }

    /// <summary>
    /// Close all opened excel file, if its not done.
    /// </summary>
    /// <param name="listExecVar"></param>
    void CloseAllOpenedExcelFile(List<ExecVar> listExecVar)
    {
        // TODO: gestion erreur!!

        foreach(ExecVar execVar in listExecVar)
        {
            if(execVar.ExecVarType== ExecVarType.ExcelFile)
                ExecInstrCloseExcelFileMgr.Exec(_excelProcessor, execVar.Value as IExcelFile);
        }
    }

    public void FireEvent(InstrBaseExecEvent execEvent)
    {
        if (EventOccurs != null)
            EventOccurs(execEvent);
    }
}
