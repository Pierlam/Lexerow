using Lexerow.Core.Diag;

namespace Lexerow.Core.System.ActivLog;

/// <summary>
/// activity logger for the Lexerow application, used in these stages/modules:
/// load, save files, script compilation and script execution.
/// </summary>
public class ActivityLogger : IActivityLogger
{
    private MessageBuilder _messageBuilder;

    public ActivityLogger(MessageBuilder messageBuilder)
    {
        _messageBuilder= messageBuilder;
    }

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
    public void LogCompilStart(ActivityLogLevel level, string operation, string param)
    {
        ActivityLog log = new ActivityLog(ActivityLogStage.Start, level, operation, param);
        log.Module = ActivityLogType.CompileScript;
        BuildMsgRaiseEvent(log);
    }

    /// <summary>
    /// stage is end, result is Ok/success.
    /// </summary>
    /// <param name="operation"></param>
    /// <param name="msg"></param>
    public void LogCompilEnd(ActivityLogLevel level, string operation, string param)
    {
        ActivityLog log = new ActivityLog(ActivityLogStage.End, level, operation, param);
        log.Module = ActivityLogType.CompileScript;
        BuildMsgRaiseEvent(log);
    }

    /// <summary>
    /// stage is end, result is Ok/success.
    /// </summary>
    /// <param name="operation"></param>
    /// <param name="msg"></param>
    public void LogCompilEnd(ActivityLogLevel level, string operation, string param, string param2)
    {
        ActivityLog log = new ActivityLog(ActivityLogStage.End, level, operation, param, param2);
        log.Module = ActivityLogType.CompileScript;
        BuildMsgRaiseEvent(log);
    }

    public void LogCompilOnGoing(ActivityLogLevel level, string operation, string param)
    {
        ActivityLog log = new ActivityLog(ActivityLogStage.OnGoing, level, operation, param);
        log.Module = ActivityLogType.CompileScript;
        BuildMsgRaiseEvent(log);
    }

    /// <summary>
    /// Error log are Important.
    /// </summary>
    /// <param name="error"></param>
    /// <param name="operation"></param>
    /// <param name="msg"></param>
    public void LogCompilEndError(ResultError error, string operation, string param)
    {
        ActivityLog log = new ActivityLog(ActivityLogStage.End, ActivityLogLevel.Info, operation, param);
        log.Module = ActivityLogType.CompileScript;
        log.Error = error;
        log.Result = ActivityLogResult.Error;
        BuildMsgRaiseEvent(log);
    }

    /// <summary>
    /// Error log are Important.
    /// </summary>
    /// <param name="error"></param>
    /// <param name="operation"></param>
    /// <param name="msg"></param>
    public void LogCompilEndError(string operation, string param)
    {
        ActivityLog log = new ActivityLog(ActivityLogStage.End, ActivityLogLevel.Info, operation, param);
        log.Module = ActivityLogType.CompileScript;
        //log.Error = error;
        log.Result = ActivityLogResult.Error;
        BuildMsgRaiseEvent(log);
    }

    /// <summary>
    /// stage is Start, result is Ok/success.
    /// </summary>
    /// <param name="operation"></param>
    /// <param name="msg"></param>
    public void LogExecStart(ActivityLogLevel level, string operation, string param)
    {
        ActivityLog log = new ActivityLog(ActivityLogStage.Start, level, operation, param);
        log.Module = ActivityLogType.ExecProg;
        BuildMsgRaiseEvent(log);
    }

    /// <summary>
    /// stage is end, result is Ok/success.
    /// </summary>
    /// <param name="operation"></param>
    /// <param name="msg"></param>
    public void LogExecEnd(ActivityLogLevel level, string operation, string param)
    {
        ActivityLog log = new ActivityLog(ActivityLogStage.End, level, operation, param);
        log.Module = ActivityLogType.ExecProg;
        BuildMsgRaiseEvent(log);
    }

    public void LogExecOnGoing(ActivityLogLevel level, string operation, string param)
    {
        ActivityLog log = new ActivityLog(ActivityLogStage.OnGoing, level, operation, param);
        log.Module = ActivityLogType.ExecProg;
        BuildMsgRaiseEvent(log);
    }

    /// <summary>
    /// Error log are Important.
    /// </summary>
    /// <param name="error"></param>
    /// <param name="operation"></param>
    /// <param name="msg"></param>
    public void LogExecEndError(ResultError error, string operation, string param)
    {
        ActivityLog log = new ActivityLog(ActivityLogStage.End, ActivityLogLevel.Info, operation, param);
        log.Module = ActivityLogType.ExecProg;
        log.Error = error;
        log.Result = ActivityLogResult.Error;
        BuildMsgRaiseEvent(log);
    }

    /// <summary>
    /// Build the log message and raise an event.
    /// </summary>
    /// <param name="log"></param>
    private void BuildMsgRaiseEvent(ActivityLog log)
    {
        if (log == null) return;

        if (!IsLogAllowed(log.Level)) return;

        log.Message = _messageBuilder.BuildMsg(log);

        ActivityLogEvent?.Invoke(this, log);
    }
}