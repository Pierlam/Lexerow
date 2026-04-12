using DocumentFormat.OpenXml.Math;
using Lexerow.Core.Diag;
using Lexerow.Core.InstrProgExec;
using Lexerow.Core.ScriptCompile;
using Lexerow.Core.ScriptLoad;
using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;
using Lexerow.Core.System.InstrDef;
using Lexerow.Core.System.ScriptDef;
using OpenExcelSdk;
using System.Diagnostics;

namespace Lexerow.Core;

/// <summary>
/// Lexerow core backend application.
/// </summary>
public class LexerowCore : IDisposable
{
    private ExcelProcessor _excelProcessor;

    private CoreData _coreData = new CoreData();

    private MessageBuilder _messageBuilder;

    private IActivityLogger _logger;

    private ScriptLoader _scriptLoader;

    private ScriptCompiler _scriptCompiler;

    private ProgramExecutor _programExecutor;

    /// <summary>
    /// Constructor.
    /// </summary>
    public LexerowCore()
    {
        // create the log/trace message builder
        _messageBuilder = new MessageBuilder();

        _logger = new ActivityLogger(_messageBuilder);

        // default activity log level: Info, only main information.
        _logger.ActiveLevel = ActivityLogLevel.Info;

        Diagnostics = new Diagnostics(_logger);

        _logger.ActivityLogEvent += logger_ActivityLogEvent;

        // main managers
        _excelProcessor = new ExcelProcessor();
        _scriptLoader = new ScriptLoader();
        _scriptCompiler = new ScriptCompiler(_logger, _coreData);
        _programExecutor = new ProgramExecutor(_logger, _excelProcessor);
    }

    public void Dispose() 
    {
        Diagnostics.CloseLogs();
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

        result= ExecuteScript(scriptName);

        Diagnostics.CloseLogs();
        return result;
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
        _logger.LogCompil(ActivityLogLevel.Info, "LexerowCore.LoadScriptFromLines", scriptName);

        Result result = new Result();

        // check that the name is not already used by another program
        if (string.IsNullOrWhiteSpace(scriptName))
        {
            if (scriptName == null) scriptName = string.Empty;
            var error = result.AddError(ErrorCode.ProgramWrongName, scriptName);
            _logger.LogCompilError(error, "LexerowCore.LoadScriptFromLines", scriptName);
            return result;
        }

        // any program should have the same name
        Program program = _coreData.GetProgramByName(scriptName);
        if (program != null)
        {
            result.AddError(ErrorCode.ProgramNameAlreadyUsed, scriptName);
            return result;
        }

        if (!_scriptLoader.LoadScriptFromLines(result, scriptName, scriptLines, out System.ScriptDef.Script script))
            return result;

        // compile the script,  generate instructions
        _scriptCompiler.CompileScript(result, script, out Program programScript);
        if (!result.Res)
            return result;

        // save it
        _coreData.ListProgram.Add(programScript);

        _logger.LogCompil(ActivityLogLevel.Info, "LoadScriptFromFile", scriptName);

        return result;
    }

    /// <summary>
    /// Load, compile and then execute a text file script.
    /// </summary>
    /// <param name="scriptname"></param>
    /// <param name="filename"></param>
    /// <returns></returns>
    public Result LoadExecScript(string scriptname, string filename)
    {
        Stopwatch stopwatch = new Stopwatch();
        stopwatch.Start();

        Result result;

        // check that the script name is set
        if (string.IsNullOrWhiteSpace(scriptname))
        {
            result=new Result();
            var error = result.AddError(ErrorCode.ScriptNameNullOrEmpty, string.Empty);
            _logger.LogCompilError(error, "LexerowCore.LoadExecScript", string.Empty);
            return result;
        }

        if (string.IsNullOrWhiteSpace(filename))
        {
            result = new Result();
            var error = result.AddError(ErrorCode.FileNameNullOrEmpty, string.Empty);
            _logger.LogCompilError(error, "LexerowCore.LoadExecScript", string.Empty);
            return result;
        }

        _logger.LogCompil(ActivityLogLevel.Info, "LexerowCore.LoadExecScript", filename);

        string elapsedTime;

        // first load the script from the file
        result = LoadScriptFromFile(scriptname, filename, out System.ScriptDef.Script script);
        if (!result.Res)
        {
            _logger.LogCompilError("LexerowCore.LoadExecScript", filename);
            return result;
        }

        // compile the script, generate instructions
        _scriptCompiler.CompileScript(result, script, out Program program);
        if (!result.Res)
        {
            _logger.LogCompilError("LexerowCore.LoadExecScript", filename);
            return result;
        }

        // save it
        _coreData.ListProgram.Add(program);

        // then, execute the compiled script
        _programExecutor.Exec(result, program);

        stopwatch.Stop();
        elapsedTime = string.Format("{0:hh\\:mm\\:ss\\.fff}", stopwatch.Elapsed);

        if (result.Res)
            _logger.LogCompil(ActivityLogLevel.Info, "LexerowCore.LoadExecScript", filename,elapsedTime);
        else
            _logger.LogCompilError("LexerowCore.LoadExecScript", filename);

        return result;
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
        _logger.LogCompil(ActivityLogLevel.Info, "LexerowCore.LoadScript", fileName);

        Result result = LoadScriptFromFile(scriptName, fileName, out System.ScriptDef.Script script);
        if (!result.Res)
            return result;

        // compile the script, generate instructions
        _scriptCompiler.CompileScript(result, script, out Program programInstr);
        if (!result.Res)
            return result;

        // save it
        _coreData.ListProgram.Add(programInstr);

        _logger.LogCompil(ActivityLogLevel.Info, "LexerowCore.LoadScript", fileName);

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
        _logger.LogExec(ActivityLogLevel.Info, "LexerowCore.ExecuteScript", scriptName);

        Result result = new Result();

        // check that the name is not already used by another program
        if (string.IsNullOrWhiteSpace(scriptName))
        {
            if (scriptName == null) scriptName = string.Empty;
            var error = result.AddError(ErrorCode.ProgramWrongName, scriptName);
            _logger.LogCompilError(error, "LexerowCore.ExecuteScript", scriptName);
            return result;
        }

        // find the program script
        Program program = _coreData.GetProgramByName(scriptName);
        if (program == null)
        {
            result.AddError(ErrorCode.ProgramNotFound, scriptName);
            //_logger.LogExecEndError(error, ErrorCode.ExecUnableOpenFile, "LexerowCore.ExecuteScript", scriptName);
            return result;
        }

        // execute ProgramScript
        _programExecutor.Exec(result, program);

        _logger.LogExec(ActivityLogLevel.Info, "LexerowCore.ExecuteScript", scriptName);
        return result;
    }

    /// <summary>
    /// Load the script form the text file.
    /// </summary>
    /// <param name="scriptName"></param>
    /// <param name="fileName"></param>
    /// <param name="script"></param>
    /// <returns></returns>
    public Result LoadScriptFromFile(string scriptName, string fileName, out System.ScriptDef.Script script)
    {
        script = null;

        _logger.LogCompil(ActivityLogLevel.Info, "LexerowCore.LoadScriptFromFile", fileName);

        Result result = new Result();

        // check that the name is not already used by another program
        if (string.IsNullOrWhiteSpace(scriptName))
        {
            if (scriptName == null) scriptName = string.Empty;
            var error = result.AddError(ErrorCode.ProgramWrongName, scriptName);
            _logger.LogCompilError(error, "LexerowCore.LoadScriptFromFile", string.Empty);
            return result;
        }

        // any program should have the same name
        Program program = _coreData.GetProgramByName(scriptName);
        if (program != null)
        {
            result.AddError(ErrorCode.ProgramNameAlreadyUsed, scriptName);
            return result;
        }

        if (!_scriptLoader.LoadScriptFromFile(result, scriptName, fileName, out script))
        {
            _logger.LogCompilError(result.ListError[0], "LexerowCore.LoadScriptFromFile", fileName);
            return result;
        }


        _logger.LogCompil(ActivityLogLevel.Info, "LexerowCore.LoadScriptFromFile", script.ScriptLines.Count.ToString());
        return result;
    }

    /// <summary>
    /// Diagnostics tool.
    /// </summary>
    public Diagnostics Diagnostics { get;}

    private void logger_ActivityLogEvent(object? sender, ActivityLog e)
    {
        // log to console is active?
        if (Diagnostics.IsLogToConsoleActive)
            Diagnostics.LogToConsole(e);

        // log to txt file is active?
        if (Diagnostics.IsSaveLogTxtActive)
            Diagnostics.SaveLog(e);

        // log outside
        ActivityLogEvent?.Invoke(sender, e);
    }
}