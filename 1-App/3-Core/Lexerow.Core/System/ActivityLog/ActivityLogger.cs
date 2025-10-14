using Lexerow.Core.System;
using Lexerow.Core.System;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System.ActivityLog;

/// <summary>
/// activity logger for the Lexerow application, used in these stages/modules:
/// load, save files, script compilation and script execution.
/// </summary>
public class ActivityLogger:IActivityLogger
{
    public event EventHandler<ActivityLogBase> ActivityLogEvent;

    /// <summary>
    /// stage is Start, result is Ok/success.
    /// </summary>
    /// <param name="operation"></param>
    /// <param name="msg"></param>
    public void LogCompilStart(ActivityLogLevel level, string operation, string msg)
    {
        ActivityLogCompileScript log = new ActivityLogCompileScript(ActivityLogStage.Start, level, operation, msg);
        RaiseEvent(log);
    }

    /// <summary>
    /// stage is end, result is Ok/success.
    /// </summary>
    /// <param name="operation"></param>
    /// <param name="msg"></param>
    public void LogCompilEnd(ActivityLogLevel level, string operation, string msg)
    {
        ActivityLogCompileScript log = new ActivityLogCompileScript(ActivityLogStage.End, level, operation, msg);
        RaiseEvent(log);
    }

    public void LogCompilOnGoing(ActivityLogLevel level, string operation, string msg)
    {
        ActivityLogCompileScript log = new ActivityLogCompileScript(ActivityLogStage.OnGoing, level, operation, msg);
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
        ActivityLogCompileScript log = new ActivityLogCompileScript(ActivityLogStage.End, ActivityLogLevel.Important, operation, msg);
        log.Error = error;
        log.Result=ActivityLogResult.Error;
        RaiseEvent(log);
    }

    public void LogExce(ActivityLogStage stage, string operation)
    { }

    public void LogLoad(ActivityLogStage stage, string operation)
    { }


    void RaiseEvent(ActivityLogBase log)
    {
        ActivityLogEvent?.Invoke(this, log);
    }
}
