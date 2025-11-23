using Lexerow.Core.ExcelLayer;
using Lexerow.Core.ProgExec;
using Lexerow.Core.ScriptCompile;
using Lexerow.Core.ScriptLoad;
using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;
using Lexerow.Core.System.Excel;
using Lexerow.Core.System.InstrDef;
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

    private ProgramExecutor _programExecutor;

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

        _programExecutor = new ProgramExecutor(_logger, _excelProcessor);
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
    public Result LoadExecLinesScript(string scriptName, List<string> scriptLines)
    {
        Result result = LoadLinesScript(scriptName, scriptLines);
        if (!result.Res) return result;

        return ExecuteScript(scriptName);
    }

    /// <summary>
    /// Load a script from lines (list of lines) and compile it.
    /// If it's ok, the script is saved in memory, ready to be executed.
    /// </summary>
    /// <param name="scriptName"></param>
    /// <param name="scriptLines"></param>
    /// <returns></returns>
    public Result LoadLinesScript(string scriptName, List<string> scriptLines)
    {
        _logger.LogCompilStart(ActivityLogLevel.Important, "LexerowCore.LoadScriptFromLines", scriptName);

        Result result = new Result();

        // check that the name is not already used by another program
        if (string.IsNullOrWhiteSpace(scriptName))
        {
            if (scriptName == null) scriptName = string.Empty;
            var error = result.AddError(ErrorCode.ProgramWrongName, scriptName);
            _logger.LogCompilEndError(error, "LexerowCore.LoadScriptFromLines", scriptName);
            return result;
        }

        // any program should have the same name
        Program program = _coreData.GetProgramByName(scriptName);
        if (program != null)
        {
            result.AddError(ErrorCode.ProgramNameAlreadyUsed, scriptName);
            return result;
        }

        if (!_scriptLoader.LoadScriptFromLines(result, scriptName, scriptLines, out Script script))
            return result;

        // compile the script,  generate instructions
        _scriptCompiler.CompileScript(result, script, out Program programScript);
        if (!result.Res)
            return result;

        // save it
        _coreData.ListProgram.Add(programScript);

        _logger.LogCompilEnd(ActivityLogLevel.Important, "LoadScriptFromFile", scriptName);

        return result;
    }

    /// <summary>
    /// Load, compile and then execute a text file script.
    /// </summary>
    /// <param name="scriptName"></param>
    /// <param name="filename"></param>
    /// <returns></returns>
    public Result LoadExecScript(string scriptName, string filename)
    {
        Result result = LoadScript(scriptName, filename);
        if (!result.Res) return result;

        return ExecuteScript(scriptName);
    }

    /// <summary>
    /// Load a script from a text file and compile it.
    /// If it's ok, the script is saved in memory, ready to be executed.
    /// </summary>
    /// <param name="progName"></param>
    /// <param name="fileName"></param>
    /// <returns></returns>
    public Result LoadScript(string scriptName, string fileName)
    {
        _logger.LogCompilStart(ActivityLogLevel.Important, "LexerowCore.LoadScriptFromFile", scriptName);

        Result result = new Result();

        // check that the name is not already used by another program
        if (string.IsNullOrWhiteSpace(scriptName))
        {
            if (scriptName == null) scriptName = string.Empty;
            var error = result.AddError(ErrorCode.ProgramWrongName, scriptName);
            _logger.LogCompilEndError(error, "LexerowCore.LoadScriptFromFile", scriptName);
            return result;
        }

        // any program should have the same name
        Program program = _coreData.GetProgramByName(scriptName);
        if (program != null)
        {
            result.AddError(ErrorCode.ProgramNameAlreadyUsed, scriptName);
            return result;
        }

        if (!_scriptLoader.LoadScriptFromFile(result, scriptName, fileName, out Script script))
            return result;

        // compile the script,  generate instructions
        _scriptCompiler.CompileScript(result, script, out Program programInstr);
        if (!result.Res)
            return result;


        // save it
        _coreData.ListProgram.Add(programInstr);

        _logger.LogCompilEnd(ActivityLogLevel.Important, "LexerowCore.LoadScriptFromFile", scriptName);

        return result;
    }

    /// <summary>
    /// Execute a script, should be compiled before.
    /// </summary>
    /// <param name="scriptName"></param>
    /// <returns></returns>
    /// <exception cref="Exception"></exception>
    public Result ExecuteScript(string scriptName)
    {
        _logger.LogExecStart(ActivityLogLevel.Important, "LexerowCore.ExecuteScript", scriptName);

        Result result = new Result();

        // check that the name is not already used by another program
        if (string.IsNullOrWhiteSpace(scriptName))
        {
            if (scriptName == null) scriptName = string.Empty;
            var error = result.AddError(ErrorCode.ProgramWrongName, scriptName);
            _logger.LogCompilEndError(error, "LexerowCore.LoadScriptFromFile", scriptName);
            return result;
        }

        // find the program script
        Program program = _coreData.GetProgramByName(scriptName);
        if (program == null)
        {
            result.AddError(ErrorCode.ProgramNotFound, scriptName);
            return result;
        }

        // execute ProgramScript
        _programExecutor.Exec(result, program);

        return result;
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