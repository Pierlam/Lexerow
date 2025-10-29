using Lexerow.Core.System;
using Lexerow.Core.System;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System.ActivLog;

/// <summary>
/// activity logger for the Lexerow application, used in these stages/modules:
/// load, save files, script compilation and script execution.
/// </summary>
public class ActivityLogger:IActivityLogger
{
    public event EventHandler<ActivityLog> ActivityLogEvent;

    /// <summary>
    /// stage is Start, result is Ok/success.
    /// </summary>
    /// <param name="operation"></param>
    /// <param name="msg"></param>
    public void LogCompilStart(ActivityLogLevel level, string operation, string msg)
    {
        ActivityLog log = new ActivityLog(ActivityLogStage.Start, level, operation, msg);
        log.ActivityLogBaseType= ActivityLogType.CompileScript;
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
        log.ActivityLogBaseType = ActivityLogType.CompileScript;
        RaiseEvent(log);
    }

    public void LogCompilOnGoing(ActivityLogLevel level, string operation, string msg)
    {
        ActivityLog log = new ActivityLog(ActivityLogStage.OnGoing, level, operation, msg);
        log.ActivityLogBaseType = ActivityLogType.CompileScript;
        RaiseEvent(log);
    }

    /// <summary>
    /// Error log are Important.
    /// </summary>
    /// <param name="error"></param>
    /// <param name="operation"></param>
    /// <param name="msg"></param>
    public void LogCompilEndError(ExecResultError error, string operation, string msg)
    {
        ActivityLog log = new ActivityLog(ActivityLogStage.End, ActivityLogLevel.Important, operation, msg);
        log.ActivityLogBaseType = ActivityLogType.CompileScript;
        log.Error = error;
        log.Result=ActivityLogResult.Error;
        RaiseEvent(log);
    }

    /// <summary>
    /// stage is Start, result is Ok/success.
    /// </summary>
    /// <param name="operation"></param>
    /// <param name="msg"></param>
    public void LogRunStart(ActivityLogLevel level, string operation, string msg)
    {
        ActivityLog log = new ActivityLog(ActivityLogStage.Start, level, operation, msg);
        log.ActivityLogBaseType = ActivityLogType.RunProg;
        RaiseEvent(log);
    }

    /// <summary>
    /// stage is end, result is Ok/success.
    /// </summary>
    /// <param name="operation"></param>
    /// <param name="msg"></param>
    public void LogRunEnd(ActivityLogLevel level, string operation, string msg)
    {
        ActivityLog log = new ActivityLog(ActivityLogStage.End, level, operation, msg);
        log.ActivityLogBaseType = ActivityLogType.RunProg;
        RaiseEvent(log);
    }

    public void LogRunOnGoing(ActivityLogLevel level, string operation, string msg)
    {
        ActivityLog log = new ActivityLog(ActivityLogStage.OnGoing, level, operation, msg);
        log.ActivityLogBaseType = ActivityLogType.RunProg;
        RaiseEvent(log);
    }

    /// <summary>
    /// Error log are Important.
    /// </summary>
    /// <param name="error"></param>
    /// <param name="operation"></param>
    /// <param name="msg"></param>
    public void LogRunEndError(ExecResultError error, string operation, string msg)
    {
        ActivityLog log = new ActivityLog(ActivityLogStage.End, ActivityLogLevel.Important, operation, msg);
        log.ActivityLogBaseType = ActivityLogType.RunProg;
        log.Error = error;
        log.Result = ActivityLogResult.Error;
        RaiseEvent(log);
    }

    void RaiseEvent(ActivityLog log)
    {
        ActivityLogEvent?.Invoke(this, log);
    }
}
