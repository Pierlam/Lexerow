namespace Lexerow.Core.System;

/// <summary>
/// Execution (load, save, compilation or execution) result error.
/// TODO: rename into ResultError??
/// </summary>
public class ExecResultError
{
    public ExecResultError(ErrorCode errorCode, string param)
    {
        DateTimeCreation = DateTime.UtcNow;
        ErrorCode = errorCode;
        Param = param;
    }

    public ExecResultError(ErrorCode errorCode, string param, string param2)
    {
        DateTimeCreation = DateTime.UtcNow;
        ErrorCode = errorCode;
        Param = param;
        Param2 = param2;
    }

    public ExecResultError(ErrorCode errorCode, Exception exception, string param)
    {
        DateTimeCreation = DateTime.UtcNow;
        ErrorCode = errorCode;
        Param = param;
    }

    public ExecResultError(ErrorCode errorCode, int lineNum, int colNum, string param)
    {
        DateTimeCreation = DateTime.UtcNow;
        ErrorCode = errorCode;
        Param = param;
        LineNum = lineNum;
        ColNum = colNum;
    }

    public ExecResultError(ErrorCode errorCode, int lineNum, int colNum, string param, string param2)
    {
        DateTimeCreation = DateTime.UtcNow;
        ErrorCode = errorCode;
        Param = param;
        Param2 = param2;
        LineNum = lineNum;
        ColNum = colNum;
    }

    public ExecResultError(ErrorCode errorCode, int lineNum, int colNum, Exception exception, string param)
    {
        DateTimeCreation = DateTime.UtcNow;
        ErrorCode = errorCode;
        Exception = exception;
        Param = param;
        LineNum = lineNum;
        ColNum = colNum;
    }

    /// <summary>
    /// When the error was created.
    /// </summary>
    public DateTime DateTimeCreation { get; private set; }

    /// <summary>
    /// The code of the error.
    /// </summary>
    public ErrorCode ErrorCode { get; set; }

    public int LineNum { get; set; } = 0;
    public int ColNum { get; set; } = 0;

    public Exception? Exception { get; set; } = null;

    /// <summary>
    /// Parameter 1 of the error.
    /// </summary>
    public string Param { get; set; } = string.Empty;

    /// <summary>
    /// Parameter 2 of the error.
    /// </summary>
    public string Param2 { get; set; } = string.Empty;
}