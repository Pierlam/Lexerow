using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System.ActivityLog;


public enum ActivityLogBaseType
{
    Other,
    LoadScript,
    CompileScript,
    ExecScript,
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

public abstract class ActivityLogBase
{
    public ActivityLogBaseType ActivityLogBaseType { get; set; }= ActivityLogBaseType.Other;    

    public DateTime When { get; private set; }= DateTime.Now;

    public ActivityLogStage Stage { get; set; } = ActivityLogStage.StartEnd;

    public ActivityLogLevel Level { get;set; }= ActivityLogLevel.Info;

    /// <summary>
    /// name of the operation/action.
    /// </summary>
    public string Operation { get; set; }=string.Empty;

    public string Msg { get; set; } = string.Empty;

    /// <summary>
    /// result of the execution of the operation.
    /// </summary>
    public ActivityLogResult Result { get; set; } = ActivityLogResult.Ok;

    public ExecResultError? Error { get; set; } = null;
}

public class ActivityLogCompileScript: ActivityLogBase
{
    public ActivityLogCompileScript(ActivityLogStage stage, ActivityLogLevel level, string operation, string msg)
    {
        ActivityLogBaseType = ActivityLogBaseType.CompileScript;
        Level = level;
        Stage = stage;
        Operation = operation;
        Msg = msg;
    }
}
