namespace Lexerow.Core.System.ActivLog;

/// <summary>
/// activity logger for the Lexerow application, used in these stages/modules:
/// load, save files, script compilation and script execution.
/// </summary>
public class ActivityLogger : IActivityLogger
{
    public event EventHandler<ActivityLog> ActivityLogEvent;

    /// <summary>
    /// Active lelve:  current allowed level for log and trace.
    /// if it's Trace, Info an Debug are allowed.
    /// If it's Info, the top one, only this one is allowed.
    /// </summary>
    public ActivityLogLevel ActiveLevel { get; set; }= ActivityLogLevel.Info;

    /// <summary>
    /// Is the provided log level allowed?
    /// exp: level is Trace and Active level is Info, -> not allowed.
    /// </summary>
    /// <param name="level"></param>
    /// <returns></returns>
    public bool IsLogAllowed(ActivityLogLevel level)
    {
        if (ActiveLevel >=level) return true;
        return false;
    }

    /// <summary>
    /// return true is the level Trace is active.
    /// </summary>
    /// <returns></returns>
    public bool IsLevelTraceActive()
    {
        if(ActiveLevel >=ActivityLogLevel.Trace)return true;
        return false;
    }

    /// <summary>
    /// return true is the level Debug (and Trace) is active.
    /// </summary>
    /// <returns></returns>
    public bool IsLevelTraceDebug()
    {
        if (ActiveLevel >= ActivityLogLevel.Debug) return true;
        return false;
    }


    /// <summary>
    /// stage is Start, result is Ok/success.
    /// </summary>
    /// <param name="operation"></param>
    /// <param name="msg"></param>
    public void LogCompilStart(ActivityLogLevel level, string operation, string msg)
    {
        ActivityLog log = new ActivityLog(ActivityLogStage.Start, level, operation, msg);
        log.Module = ActivityLogType.CompileScript;
        RaiseEvent(log);
    }

    /// <summary>
    /// stage is end, result is Ok/success.
    /// </summary>
    /// <param name="operation"></param>
    /// <param name="msg"></param>
    public void LogCompilEnd(ActivityLogLevel level, string operation, string msg)
    {
        ActivityLog log = new ActivityLog(ActivityLogStage.End, level, operation, msg);
        log.Module = ActivityLogType.CompileScript;
        RaiseEvent(log);
    }

    public void LogCompilOnGoing(ActivityLogLevel level, string operation, string msg)
    {
        ActivityLog log = new ActivityLog(ActivityLogStage.OnGoing, level, operation, msg);
        log.Module = ActivityLogType.CompileScript;
        RaiseEvent(log);
    }

    /// <summary>
    /// Error log are Important.
    /// </summary>
    /// <param name="error"></param>
    /// <param name="operation"></param>
    /// <param name="msg"></param>
    public void LogCompilEndError(ResultError error, string operation, string msg)
    {
        ActivityLog log = new ActivityLog(ActivityLogStage.End, ActivityLogLevel.Info, operation, msg);
        log.Module = ActivityLogType.CompileScript;
        log.Error = error;
        log.Result = ActivityLogResult.Error;
        RaiseEvent(log);
    }

    /// <summary>
    /// Error log are Important.
    /// </summary>
    /// <param name="error"></param>
    /// <param name="operation"></param>
    /// <param name="msg"></param>
    public void LogCompilEndError(string operation, string msg)
    {
        ActivityLog log = new ActivityLog(ActivityLogStage.End, ActivityLogLevel.Info, operation, msg);
        log.Module = ActivityLogType.CompileScript;
        //log.Error = error;
        log.Result = ActivityLogResult.Error;
        RaiseEvent(log);
    }

    /// <summary>
    /// stage is Start, result is Ok/success.
    /// </summary>
    /// <param name="operation"></param>
    /// <param name="msg"></param>
    public void LogExecStart(ActivityLogLevel level, string operation, string msg)
    {
        ActivityLog log = new ActivityLog(ActivityLogStage.Start, level, operation, msg);
        log.Module = ActivityLogType.ExecProg;
        RaiseEvent(log);
    }

    /// <summary>
    /// stage is end, result is Ok/success.
    /// </summary>
    /// <param name="operation"></param>
    /// <param name="msg"></param>
    public void LogExecEnd(ActivityLogLevel level, string operation, string msg)
    {
        ActivityLog log = new ActivityLog(ActivityLogStage.End, level, operation, msg);
        log.Module = ActivityLogType.ExecProg;
        RaiseEvent(log);
    }

    public void LogExecOnGoing(ActivityLogLevel level, string operation, string msg)
    {
        ActivityLog log = new ActivityLog(ActivityLogStage.OnGoing, level, operation, msg);
        log.Module = ActivityLogType.ExecProg;
        RaiseEvent(log);
    }

    /// <summary>
    /// Error log are Important.
    /// </summary>
    /// <param name="error"></param>
    /// <param name="operation"></param>
    /// <param name="msg"></param>
    public void LogExecEndError(ResultError error, string operation, string msg)
    {
        ActivityLog log = new ActivityLog(ActivityLogStage.End, ActivityLogLevel.Info, operation, msg);
        log.Module = ActivityLogType.ExecProg;
        log.Error = error;
        log.Result = ActivityLogResult.Error;
        RaiseEvent(log);
    }

    private void RaiseEvent(ActivityLog log)
    {
        if (log == null) return;
        if(IsLogAllowed(log.Level))
            ActivityLogEvent?.Invoke(this, log);
    }
}