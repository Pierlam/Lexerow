using Lexerow.Core.System;
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

    //public bool AddInstr(InstrBase instrBase)
    //{
    //    _coreData.ListInstr.Add(instrBase);
    //    return true;
    //}

    public ExecResult Compile()
    {
        ExecResult execResult = new ExecResult();

        if (_coreData.CurrProgramInstr == null)
        {
            execResult.AddError(new ExecResultError(ErrorCode.NoCurrentProgramExist, null));
            return execResult;
        }

        return Compile(_coreData.CurrProgramInstr);
    }

    /// <summary>
    /// Execute the current program.
    /// </summary>
    /// <returns></returns>
    public ExecResult Execute()
    {
        ExecResult execResult = new ExecResult();

        if (_coreData.CurrProgramInstr == null)
        {
            execResult.AddError(new ExecResultError(ErrorCode.NoCurrentProgramExist, null));
            return execResult;
        }

        return ExecuteProgram(_coreData.CurrProgramInstr);
    }

    /// <summary>
    /// Execute a program by the name.
    /// </summary>
    /// <param name="name"></param>
    /// <returns></returns>
    public ExecResult Execute(string name)
    {
        ExecResult execResult = new ExecResult();

        if(string.IsNullOrWhiteSpace(name))
        {
            if (name == null) name = string.Empty;
            execResult.AddError(new ExecResultError(ErrorCode.ProgramWrongName, name));
            return execResult;
        }

        ProgramInstr program= _coreData.GetProgramByName(name);
        if(program==null)
        {
            execResult.AddError(new ExecResultError(ErrorCode.ProgramNotFound, name));
            return execResult;
        }

        return ExecuteProgram(program);
    }

    ExecResult Compile(ProgramInstr program)
    {
        ExecResult execResult = new ExecResult();

        SendAppTraceCompile(AppTraceLevel.Info, "Compile all: Start");

        // possible to create the instr?
        if (program.Stage != CoreStage.Build)
        {
            execResult.AddError(new ExecResultError(ErrorCode.UnableCreateInstrNotInStageBuild, null));
            SendAppTraceCompile(AppTraceLevel.Error, ErrorCode.UnableCreateInstrNotInStageBuild.ToString());
            return execResult;
        }

        // check all instruction, one by one 
        execResult = ExecCompileInstrMgr.CheckAllInstr(program.ListInstr);
        if (!execResult.Result)
        {
            SendAppTraceCompile(AppTraceLevel.Error, "CheckAllInstr: " + execResult.ListError[0].Param);
            program.Stage = CoreStage.InstrError;
        }

        program.Stage = CoreStage.ReadyToExec;
        SendAppTraceCompile(AppTraceLevel.Info, "Compile all: End");
        return execResult;
    }

    /// <summary>
    /// Execute all saved instruction.
    /// Need to compile instr before execute them.
    /// </summary>
    /// <returns></returns>
    ExecResult ExecuteProgram(ProgramInstr program)
    {
        ExecResult execResult= new ExecResult();

        // not yet compiled?
        if (program.Stage == CoreStage.Build)
        {
            execResult = Compile();
            if (!execResult.Result)
                // error occured during the compilation process, bye
                return execResult;
        }

        program.Stage = CoreStage.Exec;

        bool res;
        // execute saved instructions, one by one
        foreach (var instr in program.ListInstr) 
        {
            if (instr.InstrType == InstrType.OpenExcel)
            {
                _execStartCurrInstr= DateTime.Now;
                res= ExecInstrOpenExcelFile(execResult, instr as InstrOpenExcel, _listExecVar, _execStartCurrInstr);
                if(!res)
                {
                    program.Stage = CoreStage.Build;
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
                res = ExecInstrForEachRowIfThen(execResult, instr as InstrOnExcelForEachRowIfThen, _listExecVar, _execStartCurrInstr);
                if (!res)
                {
                    program.Stage = CoreStage.Build;
                    return execResult;
                }
 
                continue;
            }
        }

        // close all opened excel file, if its not done
        CloseAllOpenedExcelFile(_listExecVar);

        program.Stage = CoreStage.Build;

        return execResult; 
    }

    /// <summary>
    /// Execute the instr OpenExcel.
    /// return an object: an excel file.
    /// </summary>
    /// <param name="instrOpenExcel"></param>
    /// <returns></returns>
    bool ExecInstrOpenExcelFile(ExecResult execResult, InstrOpenExcel instrOpenExcel, List<ExecVar> listExecVar, DateTime execStart)
    {
        SendAppTraceExec(AppTraceLevel.Info, "ExecInstrOpenExcelFile", InstrOpenExcelExecEvent.CreateStart(instrOpenExcel.FileName));

        if (!_excelProcessor.Open(instrOpenExcel.FileName, out IExcelFile excelFile, out ExecResultError error))
        {
            execResult.AddError(error);
            SendAppTraceExec(AppTraceLevel.Error, "ExecInstrOpenExcelFile", InstrOpenExcelExecEvent.CreateFinished(execStart, InstrBaseExecEventResult.Error, instrOpenExcel.FileName));
            return false;
        }

        ExecVar execVar = new ExecVar(instrOpenExcel.ExcelFileObjectName, ExecVarType.ExcelFile, excelFile);
        listExecVar.Add(execVar);

        SendAppTraceExec(AppTraceLevel.Info, "ExecInstrOpenExcelFile", InstrOpenExcelExecEvent.CreateFinished(execStart, InstrBaseExecEventResult.Ok, instrOpenExcel.FileName));
        return true;
    }

    /// <summary>
    /// Execute instr: 
    /// 	ForEachRowIfThen(file, sheetNum, A,Value>10, Value=10)
    /// </summary>
    /// <param name="instr"></param>
    /// <param name="listExecVar"></param>
    /// <returns></returns>
    bool ExecInstrForEachRowIfThen(ExecResult execResult, InstrOnExcelForEachRowIfThen instr, List<ExecVar> listExecVar, DateTime execStart)
    {
        SendAppTraceExec(AppTraceLevel.Info, "ExecInstrForEachRowIfThen", InstrForEachRowIfThenExecEvent.CreateStart());

        // get the file object by name
        ExecVar execVar = listExecVar.FirstOrDefault(ev => ev.Name.Equals(instr.ExcelFileObjectName, StringComparison.InvariantCultureIgnoreCase));
        if(execVar==null)
        {
            SendAppTraceExec(AppTraceLevel.Error, "ExecInstrForEachRowIfThen", InstrForEachRowIfThenExecEvent.CreateFinished(execStart, InstrBaseExecEventResult.Error));
            execResult.AddError(new ExecResultError(ErrorCode.ExcelFileObjectNameDoesNotExists, null));
            return false;
        }

        // check that the var value in an excel file object
        if (execVar.ExecVarType!= ExecVarType.ExcelFile)
        {
            // TODO: erreur the type of the var is wrong            
        }

        // execute the instr on each cols, on each datarow
        return ExecInstrForEachRowIfThenMgr.Exec(execResult, AppTraceEvent, execStart, _excelProcessor,  execVar.Value as IExcelFile, instr);
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
