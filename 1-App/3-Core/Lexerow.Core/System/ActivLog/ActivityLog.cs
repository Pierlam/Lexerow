namespace Lexerow.Core.System.ActivLog;

public enum ActivityLogType
{
    Other,
    LoadScript,
    CompileScript,
    ExecProg,
}

public enum ActivityLogLevel
{
    // error log are important
    Important,

    Info,
    Detail
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
        ActivityLogBaseType = ActivityLogType.CompileScript;
        Level = level;
        Stage = stage;
        Operation = operation;
        Msg = msg;
    }

    public ActivityLogType ActivityLogBaseType { get; set; } = ActivityLogType.Other;

    public DateTime When { get; private set; } = DateTime.Now;

    public ActivityLogStage Stage { get; set; } = ActivityLogStage.StartEnd;

    public ActivityLogLevel Level { get; set; } = ActivityLogLevel.Info;

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