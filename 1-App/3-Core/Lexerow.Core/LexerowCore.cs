using Lexerow.Core.ExcelLayer;
using Lexerow.Core.ProgExec;
using Lexerow.Core.ScriptCompile;
using Lexerow.Core.ScriptLoad;
using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;
using Lexerow.Core.System.Excel;
using Lexerow.Core.System.ScriptDef;

namespace Lexerow.Core;

/// <summary>
/// Lexerow core backend application.
/// </summary>
public class LexerowCore
{
    private IExcelProcessor _excelProcessor;

    private CoreData _coreData = new CoreData();

    private IActivityLogger _logger;

    private ScriptLoader _scriptLoader;

    private ScriptCompiler _scriptCompiler;

    private ProgramExecutor _programRunner;

    /// <summary>
    /// Constructor.
    /// </summary>
    public LexerowCore()
    {
        _logger = new ActivityLogger();
        _logger.ActivityLogEvent += logger_ActivityLogEvent;

        _excelProcessor = new ExcelProcessorNpoi();

        _scriptLoader = new ScriptLoader();
        _scriptCompiler = new ScriptCompiler(_logger, _coreData);

        _programRunner = new ProgramExecutor(_logger, _excelProcessor);
    }

    public event EventHandler<ActivityLog> ActivityLogEvent;

    /// <summary>
    /// Execute script program.
    /// </summary>
    //Exec Exec { get; set; }

    /// <summary>
    /// Load, compile and then execute a lines script.
    /// </summary>
    /// <param name="scriptName"></param>
    /// <param name="filename"></param>
    /// <returns></returns>
    public ExecResult LoadExecLinesScript(string scriptName, List<string> scriptLines)
    {
        ExecResult execResult = LoadLinesScript(scriptName, scriptLines);
        if (!execResult.Result) return execResult;

        return ExecuteScript(scriptName);
    }

    /// <summary>
    /// Load a script from lines (list of lines) and compile it.
    /// If it's ok, the script is saved in memory, ready to be executed.
    /// </summary>
    /// <param name="scriptName"></param>
    /// <param name="scriptLines"></param>
    /// <returns></returns>
    public ExecResult LoadLinesScript(string scriptName, List<string> scriptLines)
    {
        _logger.LogCompilStart(ActivityLogLevel.Important, "LexerowCore.LoadScriptFromLines", scriptName);

        ExecResult execResult = new ExecResult();

        // check that the name is not already used by another program
        if (string.IsNullOrWhiteSpace(scriptName))
        {
            if (scriptName == null) scriptName = string.Empty;
            var error = execResult.AddError(ErrorCode.ProgramWrongName, scriptName);
            _logger.LogCompilEndError(error, "LexerowCore.LoadScriptFromLines", scriptName);
            return execResult;
        }

        // any program should have the same name
        ProgramScript program = _coreData.GetProgramByName(scriptName);
        if (program != null)
        {
            execResult.AddError(ErrorCode.ProgramNameAlreadyUsed, scriptName);
            return execResult;
        }

        if (!_scriptLoader.LoadScriptFromLines(execResult, scriptName, scriptLines, out Script script))
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

    /// <summary>
    /// Load, compile and then execute a text file script.
    /// </summary>
    /// <param name="scriptName"></param>
    /// <param name="filename"></param>
    /// <returns></returns>
    public ExecResult LoadExecScript(string scriptName, string filename)
    {
        ExecResult execResult = LoadScript(scriptName, filename);
        if (!execResult.Result) return execResult;

        return ExecuteScript(scriptName);
    }

    /// <summary>
    /// Load a script from a text file and compile it.
    /// If it's ok, the script is saved in memory, ready to be executed.
    /// </summary>
    /// <param name="progName"></param>
    /// <param name="fileName"></param>
    /// <returns></returns>
    public ExecResult LoadScript(string scriptName, string fileName)
    {
        _logger.LogCompilStart(ActivityLogLevel.Important, "LexerowCore.LoadScriptFromFile", scriptName);

        ExecResult execResult = new ExecResult();

        // check that the name is not already used by another program
        if (string.IsNullOrWhiteSpace(scriptName))
        {
            if (scriptName == null) scriptName = string.Empty;
            var error = execResult.AddError(ErrorCode.ProgramWrongName, scriptName);
            _logger.LogCompilEndError(error, "LexerowCore.LoadScriptFromFile", scriptName);
            return execResult;
        }

        // any program should have the same name
        ProgramScript program = _coreData.GetProgramByName(scriptName);
        if (program != null)
        {
            execResult.AddError(ErrorCode.ProgramNameAlreadyUsed, scriptName);
            return execResult;
        }

        if (!_scriptLoader.LoadScriptFromFile(execResult, scriptName, fileName, out Script script))
            return execResult;

        // compile the script,  generate instructions
        _scriptCompiler.CompileScript(execResult, script, out List<InstrBase> listInstr);
        if (!execResult.Result)
            return execResult;

        // create the program
        ProgramScript programInstr = new ProgramScript(script, listInstr);

        // save it
        _coreData.ListProgram.Add(programInstr);

        _logger.LogCompilEnd(ActivityLogLevel.Important, "LexerowCore.LoadScriptFromFile", scriptName);

        return execResult;
    }

    /// <summary>
    /// Execute a script, should be compiled before.
    /// </summary>
    /// <param name="scriptName"></param>
    /// <returns></returns>
    /// <exception cref="Exception"></exception>
    public ExecResult ExecuteScript(string scriptName)
    {
        _logger.LogExecStart(ActivityLogLevel.Important, "LexerowCore.ExecuteScript", scriptName);

        ExecResult execResult = new ExecResult();

        // check that the name is not already used by another program
        if (string.IsNullOrWhiteSpace(scriptName))
        {
            if (scriptName == null) scriptName = string.Empty;
            var error = execResult.AddError(ErrorCode.ProgramWrongName, scriptName);
            _logger.LogCompilEndError(error, "LexerowCore.LoadScriptFromFile", scriptName);
            return execResult;
        }

        // find the program script
        ProgramScript program = _coreData.GetProgramByName(scriptName);
        if (program == null)
        {
            execResult.AddError(ErrorCode.ProgramNotFound, scriptName);
            return execResult;
        }

        // execute ProgramScript
        _programRunner.Exec(execResult, program);

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