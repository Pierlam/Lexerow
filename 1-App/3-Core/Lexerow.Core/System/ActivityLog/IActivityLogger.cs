using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System.ActivityLog;
public interface IActivityLogger
{
    event EventHandler<ActivityLogBase> ActivityLogEvent;

    void LogCompilStart(ActivityLogLevel level, string operation, string msg);

    /// <summary>
    /// stage is end, result is Ok/success.
    /// </summary>
    /// <param name="operation"></param>
    /// <param name="msg"></param>
    void LogCompilEnd(ActivityLogLevel level, string operation, string msg);

    void LogCompilOnGoing(ActivityLogLevel level, string operation, string msg);

    void LogCompilEndError(ExecResultError error, string operation, string msg);

    void LogExce(ActivityLogStage stage, string operation);

    void LogLoad(ActivityLogStage stage, string operation);

}
