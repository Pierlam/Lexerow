namespace Lexerow.Core.System.ActivLog;

public interface IActivityLogger
{
    event EventHandler<ActivityLog> ActivityLogEvent;

    /// <summary>
    /// Active lelve:  current allowed level for log and trace.
    /// if it's Trace, Info an Debug are allowed.
    /// If it's Info, the top one, only this one is allowed.
    /// </summary>
    ActivityLogLevel ActiveLevel { get; set; }

    /// <summary>
    /// Is the provided log level allowed?
    /// exp: level is Trace and Active level is Info, -> not allowed.
    /// </summary>
    /// <param name="level"></param>
    /// <returns></returns>
    bool IsLogAllowed(ActivityLogLevel level);

    /// <summary>
    /// return true is the level Trace is active.
    /// </summary>
    /// <returns></returns>
    bool IsLevelTraceActive();

    /// <summary>
    /// return true is the level Debug (and Trace) is active.
    /// </summary>
    /// <returns></returns>
    bool IsLevelTraceDebug();

    void LogCompilStart(ActivityLogLevel level, string operation, string param);

    /// <summary>
    /// stage is end, result is Ok/success.
    /// </summary>
    /// <param name="operation"></param>
    /// <param name="msg"></param>
    void LogCompilEnd(ActivityLogLevel level, string operation, string param);

    void LogCompilEnd(ActivityLogLevel level, string operation, string param, string param2);

    void LogCompilOnGoing(ActivityLogLevel level, string operation, string param);

    void LogCompilEndError(ResultError error, string operation, string param);

    void LogCompilEndError(string operation, string param);

    void LogExecStart(ActivityLogLevel level, string operation, string param);

    void LogExecEnd(ActivityLogLevel level, string operation, string param);

    void LogExecOnGoing(ActivityLogLevel level, string operation, string param);

    void LogExecEndError(ResultError error, string operation, string param);
}