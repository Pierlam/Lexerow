using Lexerow.Core.ExcelLayer;
using Lexerow.Core.System;
using Lexerow.Core.System.Compilator;
using Lexerow.Core.System.Excel;
using NPOI.HPSF;
using NPOI.SS.Formula.Functions;
using Lexerow.Core.System.ActivLog;
using Lexerow.Core.ScriptCompile;
using Lexerow.Core.ScriptLoad;
using Lexerow.Core.Core.Exec;
using Lexerow.Core.ProgRun;

namespace Lexerow.Core;

/// <summary>
/// Lexerow core backend application.
/// </summary>
public class LexerowCore
{
    IExcelProcessor _excelProcessor;

    CoreData _coreData=new CoreData();

    IActivityLogger _logger;

    ScriptLoader _scriptLoader;

    ScriptCompiler _scriptCompiler;

    ProgramRunner _programRunner;

    /// <summary>
    /// Constructor.
    /// </summary>
    public LexerowCore()
    {
        _logger= new ActivityLogger();
        _logger.ActivityLogEvent += logger_ActivityLogEvent;

        _excelProcessor = new ExcelProcessorNpoi();

        _scriptLoader= new ScriptLoader();
        _scriptCompiler = new ScriptCompiler(_logger, _coreData);

        _programRunner = new ProgramRunner(_logger, _excelProcessor);

        // TODO: will be removed!
        Exec = new Exec(_coreData, _excelProcessor);

    }

    public event EventHandler<ActivityLog> ActivityLogEvent;

    /// <summary>
    /// Execute script program.
    /// </summary>
    Exec Exec { get; set; }

    /// <summary>
    /// Load a script from a text file and compile it.
    /// </summary>
    /// <param name="progName"></param>
    /// <param name="fileName"></param>
    /// <returns></returns>
    public ExecResult LoadScriptFromFile(string scriptName, string fileName)
    {
        _logger.LogCompilStart(ActivityLogLevel.Important, "LoadScriptFromFile", scriptName);

        ExecResult execResult = new ExecResult();

        // check that the name is not already used by another program
        if (string.IsNullOrWhiteSpace(scriptName))
        {
            if (scriptName == null) scriptName = string.Empty;
            var error= execResult.AddError(ErrorCode.ProgramWrongName, scriptName);
            _logger.LogCompilEndError(error, "LoadScriptFromFile", scriptName);
            return execResult;
        }

        // any program should have the same name
        ProgramScript program = _coreData.GetProgramByName(scriptName);
        if (program != null)
        {
            execResult.AddError(ErrorCode.ProgramNameAlreadyUsed, scriptName);
            return execResult;
        }

        _scriptLoader.LoadScriptFromFile(execResult, scriptName, fileName, out Script script);
        if(!execResult.Result)
            return execResult;

        // compile the script,  generate instructions
        _scriptCompiler.CompileScript(execResult, script, out List<InstrBase> listInstr);
        if (!execResult.Result)
            return execResult;

        // create the program
        ProgramScript programInstr = new ProgramScript(script, listInstr);

        // save it
        _coreData.ListProgram.Add(programInstr);

        _logger.LogCompilEnd(ActivityLogLevel.Important, "LoadScriptFromFile", scriptName);

        return execResult;
    }

    public ExecResult LoadScriptFromLines(string scriptName, List<string> scriptLines)
    {
        // TODO:
        return null;
    }

    //ExecResult LoadExecScriptFromFile(string name, string filename)

    //ExecResult LoadExecScriptFromLines(string name, List<string> scriptLines)


    /// <summary>
    /// Execute a script, should be compiled before.
    /// </summary>
    /// <param name="scriptName"></param>
    /// <returns></returns>
    /// <exception cref="Exception"></exception>
    public ExecResult ExecuteScript(string scriptName)
    {
        _logger.LogRunStart(ActivityLogLevel.Important, "ExecuteScript", scriptName);

        ExecResult execResult = new ExecResult();

        // check that the name is not already used by another program
        if (string.IsNullOrWhiteSpace(scriptName))
        {
            if (scriptName == null) scriptName = string.Empty;
            var error = execResult.AddError(ErrorCode.ProgramWrongName, scriptName);
            // TODO: pb de logger!!!
            _logger.LogCompilEndError(error, "LoadScriptFromFile", scriptName);
            return execResult;
        }

        // any program should have the same name
        ProgramScript program = _coreData.GetProgramByName(scriptName);
        if (program != null)
        {
            execResult.AddError(ErrorCode.ProgramNameAlreadyUsed, scriptName);
            return execResult;
        }

        // execute ProgramScript
        _programRunner.Run(execResult, program);

        return execResult;
    }


    private void logger_ActivityLogEvent(object? sender, ActivityLog e)
    {
        ActivityLogEvent?.Invoke(sender, e);
    }



    //Action<AppTrace> _appTraceEvent;

    //public Action<AppTrace> AppTraceEvent 
    //{
    //    get { return _appTraceEvent; } 
    //    set { 
    //        //ProgBuilder.AppTraceEvent = value;
    //        Exec.AppTraceEvent = value;
    //    }
    //}
}
