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
using System.Xml.Linq;

namespace Lexerow.Core;

/// <summary>
/// Execute instructions.
/// </summary>
public class Exec
{

    IExcelProcessor _excelProcessor;

    ExecFunctionMgr _execFunctionMgr;

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
        _excelProcessor = excelProcessor;
        _execFunctionMgr = new ExecFunctionMgr(excelProcessor);
        _coreData = coreData;

    }

    //public Action<AppTrace> AppTraceEvent { get; set; }
    Action<AppTrace> _appTraceEvent;

    public Action<AppTrace> AppTraceEvent
    {
        get { return _appTraceEvent; }
        set
        {
            _appTraceEvent = value;
            _execFunctionMgr.AppTraceEvent = value;
        }
    }


    public ExecResult CompileProgram()
    {
        ExecResult execResult = new ExecResult();

        if (_coreData.CurrProgramScript == null)
        {
            execResult.AddError(new ExecResultError(ErrorCode.NoCurrentProgramExist, null));
            return execResult;
        }

        return CompileProgram(_coreData.CurrProgramScript);
    }

    //public ExecResult CompileProgram(string programName)
    //{
    //    ExecResult execResult = new ExecResult();

    //    if (_coreData.CurrProgramInstr == null)
    //    {
    //        execResult.AddError(new ExecResultError(ErrorCode.NoCurrentProgramExist, null));
    //        return execResult;
    //    }

    //    if (string.IsNullOrWhiteSpace(programName))
    //    {
    //        if (programName == null) programName = string.Empty;
    //        execResult.AddError(new ExecResultError(ErrorCode.ProgramWrongName, programName));
    //        return execResult;
    //    }

    //    ProgramInstr program = _coreData.GetProgramByName(programName);
    //    if (program == null)
    //    {
    //        execResult.AddError(new ExecResultError(ErrorCode.ProgramNotFound, programName));
    //        return execResult;
    //    }

    //    return CompileProgram(_coreData.CurrProgramInstr);
    //}

    /// <summary>
    /// Execute the current program.
    /// </summary>
    /// <returns></returns>
    public ExecResult ExecuteProgram()
    {
        ExecResult execResult = new ExecResult();

        if (_coreData.CurrProgramScript == null)
        {
            execResult.AddError(new ExecResultError(ErrorCode.NoCurrentProgramExist, null));
            return execResult;
        }

        var res= ExecuteProgram(_coreData.CurrProgramScript, _listExecVar);
        _coreData.CurrProgramScript.Stage = CoreStage.Build;
        return res;
    }

    /// <summary>
    /// Execute a program by the name.
    /// </summary>
    /// <param name="name"></param>
    /// <returns></returns>
    public ExecResult ExecuteProgram(string name)
    {
        ExecResult execResult = new ExecResult();

        if(string.IsNullOrWhiteSpace(name))
        {
            if (name == null) name = string.Empty;
            execResult.AddError(new ExecResultError(ErrorCode.ProgramWrongName, name));
            return execResult;
        }

        ProgramScript program= _coreData.GetProgramByName(name);
        if(program==null)
        {
            execResult.AddError(new ExecResultError(ErrorCode.ProgramNotFound, name));
            return execResult;
        }

        var res = ExecuteProgram(program, _listExecVar);
        _coreData.CurrProgramScript.Stage = CoreStage.Build;
        return res;
    }

    ExecResult CompileProgram(ProgramScript program)
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
    ExecResult ExecuteProgram(ProgramScript program, List<ExecVar> listExecVar)
    {
        ExecResult execResult= new ExecResult();

        // not yet compiled?
        if (program.Stage == CoreStage.Build)
        {
            execResult = CompileProgram();
            if (!execResult.Result)
                // error occured during the compilation process, bye
                return execResult;
        }

        // create a stack
        Stack<ExecTokBase> stackInstr = new Stack<ExecTokBase>();

        program.Stage = CoreStage.Exec;

        bool res;
        // execute saved instructions, one by one
        foreach (var instr in program.ListInstr) 
        {
            _execStartCurrInstr = DateTime.Now;

            //--end of line reached
            if (instr.ExecTokType == ExecTokType.Eol)
            {
                ExecStackedInstr(execResult, stackInstr, listExecVar);
                continue;
            }

            //--open bracket
            if (instr.ExecTokType == ExecTokType.OpenBracket)
            {
                // just push the instr on the stack
                stackInstr.Push(instr);
                continue;
            }

            //--param separator ,  
            // TODO: 

            //--close bracket, end of function parameters
            if (instr.ExecTokType == ExecTokType.CloseBracket)
            {
                if (!ExecCloseBracketReached(execResult, stackInstr, _listExecVar, _execStartCurrInstr))
                    return execResult;
                continue;
            }

            //--const value
            if (instr.ExecTokType == ExecTokType.ConstValue)
            {
                // just push the instr on the stack
                stackInstr.Push(instr);
                continue;
            }

            //--set var
            if (instr.ExecTokType == ExecTokType.SetVar)
            {
                // checks with instr already saved in the stack
                // TODO: don't know 

                // just push the instr on the stack
                stackInstr.Push(instr);
                continue;
            }

            if (instr.ExecTokType == ExecTokType.OpenExcel)
            {
                // just push the instr on the stack
                stackInstr.Push(instr);
                continue;
            }

            if(instr.ExecTokType == ExecTokType.CloseExcel)
            {
                // TODO: ExecInstrCloseExcelFile
                continue;
            }

            if(instr.ExecTokType == ExecTokType.ForEachRowIfThen)
            {
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

        // check the stack, should be empty
        if (stackInstr.Count > 0) 
            // TODO: set a right error code!  
            execResult.AddError(new ExecResultError(ErrorCode.InstrNotExpected, null));

        program.Stage = CoreStage.Build;

        return execResult; 
    }

    /// <summary>
    /// End of line reached.
    /// Execute stacked instructions.
    /// </summary>
    /// <param name="execResult"></param>
    /// <param name="stackInstr"></param>
    /// <returns></returns>
    bool ExecStackedInstr(ExecResult execResult, Stack<ExecTokBase> stackInstr, List<ExecVar> listExecVar)
    {
        ExecTokBase instrTop= stackInstr.Peek();

        //--SetVar excelFileObj ?
        InstrExcelFileObject instrExcelFileObject = instrTop as InstrExcelFileObject;
        if (instrExcelFileObject!= null)
        {
            // remove the obj from the stack
            stackInstr.Pop();

            // check the stack content
            // TODO:

            // the previous one is SetVar?
            InstrSetVar instrSetVar = stackInstr.Peek() as InstrSetVar;
            if (instrSetVar == null) 
            {
                // error!
                throw new Exception("todo: error");
            }

            // remove the obj from the stack
            stackInstr.Pop();

            //--it's a SetVar
            ExecVar execVar = new ExecVar(instrSetVar.VarName, ExecVarType.ExcelFile, instrExcelFileObject);
            listExecVar.Add(execVar);
            return true;
        }

        //--others cases!
        // TODO:

        return true;
    }

    /// <summary>
    /// close bracket reached, end of function parameters.
    /// 3 cases: 
    ///     function has no   parameter:  fct()
    ///     function has only parameter:  fct(param)
    ///     function has many parameters: fct(p1, p2, p3)
    ///     
    /// The close bracket is not saved in the stack    
    /// </summary>
    /// <param name="execResult"></param>
    /// <param name="stackInstr"></param>
    /// <returns></returns>
    bool ExecCloseBracketReached(ExecResult execResult, Stack<ExecTokBase> stackInstr, List<ExecVar> listExecVar, DateTime execStart)
    {
        // the stack should contains 2 item at least
        if(stackInstr.Count < 2)
        {
            // TODO: set a right error code!  
            execResult.AddError(new ExecResultError(ErrorCode.FileNameNotFound, null));
            return false;
        }

        List<ExecTokBase> listFuncParams = new List<ExecTokBase>();

        // read the item on the top of the stack, the last added
        ExecTokBase instrBaseLast = stackInstr.Peek();

        //--case 1: no parameter? the prev item is an open bracket
        if (instrBaseLast.ExecTokType == ExecTokType.OpenBracket)
        {
            // remove the open bracket
            stackInstr.Pop();

            // get the function 
            instrBaseLast = stackInstr.Pop();

            // execute the function, has no parameter
            return _execFunctionMgr.ExecFunction(execResult, stackInstr, instrBaseLast, listFuncParams, listExecVar, execStart);
        }

        // the stack should contains 3 item at least
        if (stackInstr.Count < 3)
        {
            // TODO: set a right error code!  
            execResult.AddError(new ExecResultError(ErrorCode.FileNameNotFound, null));
            return false;
        }

        // get the top of the stack
        instrBaseLast = stackInstr.Pop();

        // read the next one
        ExecTokBase instrBasePrev = stackInstr.Peek();

        //--case 2: one parameter?  fct(param
        // TODO: the prev item should be a const value, the prev-prev should be an open bracket
        if (instrBasePrev.ExecTokType == ExecTokType.OpenBracket) 
        {
            // so the prev item is a param, it should be a const value
            if (instrBaseLast.ExecTokType != ExecTokType.ConstValue)
            {
                // TODO: set a right error code!  
                execResult.AddError(new ExecResultError(ErrorCode.FileNameNotFound, null));
                return false;
            }

            // remove the open brack from the stack
            instrBasePrev = stackInstr.Pop();

            // get the function from the stack, provide the parameter
            instrBasePrev = stackInstr.Pop();

            // execute the function
            listFuncParams.Add(instrBaseLast);
            return _execFunctionMgr.ExecFunction(execResult, stackInstr, instrBasePrev, listFuncParams, listExecVar, execStart);
        }

        //--case 3: more parameter?
        // TODO: the prev-prev item is a comma param separator

        return true;
    }


    /// <summary>
    /// Execute instr: 
    /// 	ForEachRowIfThen(file, sheetNum, A,Value>10, Value=10)
    /// 	
    /// TODO: REWORK IT
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
                ExecInstrCloseExcelFileMgr.Exec(_excelProcessor, (execVar.Value as InstrExcelFileObject).ExcelFile);
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
