namespace Lexerow.Core.System.ActivLog;

public interface IActivityLogger
{
    event EventHandler<ActivityLog> ActivityLogEvent;

    /// <summary>
    /// Active level:  current allowed level for log and trace.
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

    /// <summary>
    /// stage is end, result is Ok/success.
    /// </summary>
    /// <param name="operation"></param>
    /// <param name="msg"></param>
    void Log(ActivityLogLevel level, string operation, string param);

    void Log(ActivityLogLevel level, string operation, string param, string param2);

    void LogError(string operation, ResultError error);

    void LogError(string operation, string param);

    void LogError(string operation, string param, string param2);


    void LogWarning(string operation, ResultError error);

    void LogWarning(string operation, string param, string param2);
}