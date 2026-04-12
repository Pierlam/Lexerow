namespace Lexerow.Core.System.ActivLog;

/// <summary>
/// Module
/// </summary>
public enum ActivityLogType
{
    Other,
    LoadScript,
    CompileScript,
    ExecProg,
}

public enum ActivityLogLevel
{
    /// <summary>
    /// stop the log activiy.
    /// </summary>
    Off=0, 

    /// <summary>
    ///  Generally useful information to log (service start/stop, configuration assumptions, etc). 
    ///  Info I want to always have available but usually don't care about under normal circumstances.
    /// </summary>
    Info=1,

    /// <summary>
    /// Information that is diagnostically helpful to people more than just developers (IT, sysadmins, etc.).
    /// </summary>
    Debug=2,

    /// <summary>
    ///  Only to "trace" the code and trying to find one part of a function specifically.
    /// </summary>
    Trace=3
}   

public enum ActivityLogResult
{
    Ok,
    Error,
    Warning
}

public class ActivityLog
{
    public ActivityLog(ActivityLogLevel level, string operation, string param)
    {
        Module = ActivityLogType.CompileScript;
        Level = level;
        Operation = operation;
        Param = param;
    }

    public ActivityLog(ActivityLogLevel level, string operation, string param, string param2)
    {
        Module = ActivityLogType.CompileScript;
        Level = level;
        Operation = operation;
        Param = param;
        Param2 = param2;
    }

    public ActivityLogType Module { get; set; } = ActivityLogType.Other;

    public DateTime When { get; private set; } = DateTime.Now;

    public ActivityLogLevel Level { get; set; } = ActivityLogLevel.Debug;

    /// <summary>
    /// name of the operation/action.
    /// </summary>
    public string Operation { get; set; } = string.Empty;

    /// <summary>
    /// Parameter of the log/trace, exp: file name, script name, etc.
    /// </summary>
    public string Param { get; set; } = string.Empty;

    public string Param2 { get; set; } = string.Empty;

    /// <summary>
    /// result of the execution of the operation.
    /// </summary>
    public ActivityLogResult Result { get; set; } = ActivityLogResult.Ok;

    public ResultError? Error { get; set; } = null;

    /// <summary>
    /// Human readable message.
    /// builded by the ActivityLogger, based on the registry of messages and the parameters of the log.
    /// </summary>
    public string Message { get; set; } = string.Empty;
}