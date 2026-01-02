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
    ///  Only when I would be "tracing" the code and trying to find one part of a function specifically.
    /// </summary>
    Trace=3
}   

public enum ActivityLogStage
{
    StartEnd,
    Start,
    OnGoing,
    End
}

public enum ActivityLogResult
{
    Ok,
    Error,
    Warning
}

public class ActivityLog
{
    public ActivityLog(ActivityLogStage stage, ActivityLogLevel level, string operation, string msg)
    {
        Module = ActivityLogType.CompileScript;
        Level = level;
        Stage = stage;
        Operation = operation;
        Msg = msg;
    }

    public ActivityLogType Module { get; set; } = ActivityLogType.Other;

    public DateTime When { get; private set; } = DateTime.Now;

    public ActivityLogStage Stage { get; set; } = ActivityLogStage.StartEnd;

    public ActivityLogLevel Level { get; set; } = ActivityLogLevel.Debug;

    /// <summary>
    /// name of the operation/action.
    /// </summary>
    public string Operation { get; set; } = string.Empty;

    public string Msg { get; set; } = string.Empty;

    /// <summary>
    /// result of the execution of the operation.
    /// </summary>
    public ActivityLogResult Result { get; set; } = ActivityLogResult.Ok;

    public ResultError? Error { get; set; } = null;
}