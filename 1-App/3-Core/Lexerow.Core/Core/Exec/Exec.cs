﻿using Lexerow.Core.System;
using Lexerow.Core.System.Excel;
using Lexerow.Core.System.Exec.Event;
using Lexerow.Core.Utils;
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
/// Execute instructions.
/// </summary>
public class Exec
{

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
    public Exec(CoreData coreData, IExcelProcessor excelProcessor)
    {
        _coreData = coreData;
        _excelProcessor = excelProcessor;
    }

    public Action<AppTrace> AppTraceEvent { get; set; }

    public bool AddInstr(InstrBase instrBase)
    {
        _coreData.ListInstr.Add(instrBase);
        return true;
    }

    public ExecResult Compile()
    {
        ExecResult execResult = new ExecResult();

        SendAppTraceCompile(AppTraceLevel.Info, "Compile all: Start");

        // possible to create the instr?
        if (_coreData.Stage != CoreStage.Build)
        {
            execResult.AddError(new CoreError(ErrorCode.UnableCreateInstrNotInStageBuild, null));
            SendAppTraceCompile(AppTraceLevel.Error, ErrorCode.UnableCreateInstrNotInStageBuild.ToString());
            return execResult;
        }

        // check all instruction, one by one 
        execResult= ExecCompileInstrMgr.CheckAllInstr(_coreData.ListInstr);
        if(!execResult.Result)
        {
            SendAppTraceCompile(AppTraceLevel.Error, "CheckAllInstr: " + execResult.ListError[0].Param);
            _coreData.Stage = CoreStage.InstrError;
        }

        _coreData.Stage = CoreStage.ReadyToExec;
        SendAppTraceCompile(AppTraceLevel.Info, "Compile all: End");
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

        // not yet compiled?
        if (_coreData.Stage == CoreStage.Build)
        {
            execResult = Compile();
            if (!execResult.Result)
                // error occured during the compilation process, bye
                return execResult;
        }

        _coreData.Stage = CoreStage.Exec;

        // execute saved instructions, one by one
        foreach (var instr in _coreData.ListInstr) 
        {
            if (instr.InstrType == InstrType.OpenExcel)
            {
                _execStartCurrInstr= DateTime.Now;
                execResult= ExecInstrOpenExcelFile(instr as InstrOpenExcel, _listExecVar, _execStartCurrInstr);
                if(!execResult.Result)
                {
                    _coreData.Stage = CoreStage.Build;
                    return execResult;
                }

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
                execResult = ExecInstrForEachRowIfThen(instr as InstrOnExcelForEachRowIfThen, _listExecVar, _execStartCurrInstr);
                if (!execResult.Result)
                {
                    _coreData.Stage = CoreStage.Build;
                    return execResult;
                }
 
                continue;
            }
        }

        // close all opened excel file, if its not done
        CloseAllOpenedExcelFile(_listExecVar);

        _coreData.Stage = CoreStage.Build;

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
        SendAppTraceExec(AppTraceLevel.Info, "ExecInstrOpenExcelFile", InstrOpenExcelExecEvent.CreateStart(instrOpenExcel.FileName));

        if (!_excelProcessor.Open(instrOpenExcel.FileName, out IExcelFile excelFile, out CoreError error))
        {
            result.AddError(error);
            SendAppTraceExec(AppTraceLevel.Error, "ExecInstrOpenExcelFile", InstrOpenExcelExecEvent.CreateFinished(execStart, InstrBaseExecEventResult.Error, instrOpenExcel.FileName));
            return result;
        }

        ExecVar execVar = new ExecVar(instrOpenExcel.ExcelFileObjectName, ExecVarType.ExcelFile, excelFile);
        listExecVar.Add(execVar);

        SendAppTraceExec(AppTraceLevel.Info, "ExecInstrOpenExcelFile", InstrOpenExcelExecEvent.CreateFinished(execStart, InstrBaseExecEventResult.Ok, instrOpenExcel.FileName));
        return result;
    }

    /// <summary>
    /// Execute instr: 
    /// 	ForEachRowIfThen(file, sheetNum, A,Value>10, Value=10)
    /// </summary>
    /// <param name="instr"></param>
    /// <param name="listExecVar"></param>
    /// <returns></returns>
    ExecResult ExecInstrForEachRowIfThen(InstrOnExcelForEachRowIfThen instr, List<ExecVar> listExecVar, DateTime execStart)
    {
        ExecResult execResult = new ExecResult();
        SendAppTraceExec(AppTraceLevel.Info, "ExecInstrForEachRowIfThen", InstrForEachRowIfThenExecEvent.CreateStart());

        // get the file object by name
        ExecVar execVar = listExecVar.FirstOrDefault(ev => ev.Name.Equals(instr.ExcelFileObjectName, StringComparison.InvariantCultureIgnoreCase));
        if(execVar==null)
        {
            SendAppTraceExec(AppTraceLevel.Error, "ExecInstrForEachRowIfThen", InstrForEachRowIfThenExecEvent.CreateFinished(execStart, InstrBaseExecEventResult.Error));
            execResult.AddError(new CoreError(ErrorCode.ExcelFileObjectNameDoesNotExists, null));
            return execResult;
        }

        // check that the var value in an excel file object
        if (execVar.ExecVarType!= ExecVarType.ExcelFile)
        {
            // TODO: erreur the type of the var is wrong            
        }

        // execute the instr on each cols, on each datarow
        execResult = ExecInstrForEachRowIfThenMgr.Exec(AppTraceEvent, execStart, _excelProcessor,  execVar.Value as IExcelFile, instr);

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

    public void SendAppTraceCompile(AppTraceLevel level, string msg)
    {
        if (AppTraceEvent == null) return;

        AppTrace appTrace = new AppTrace(AppTraceModule.Compile, level, msg);
        AppTraceEvent(appTrace);
    }

    public void SendAppTraceExec(AppTraceLevel level, string msg, InstrBaseExecEvent execEvent)
    {
        if (AppTraceEvent == null) return;

        AppTrace appTrace = new AppTrace(AppTraceModule.Exec, level, msg, execEvent);
        AppTraceEvent(appTrace);
    }

}
