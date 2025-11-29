using System.Diagnostics.Metrics;

namespace Lexerow.Core.System;

/// <summary>
/// Execution (load, save, compilation or execution) result error.
/// </summary>
public class ResultError
{
    public ResultError(ErrorCode errorCode, string param)
    {
        DateTimeCreation = DateTime.UtcNow;
        ErrorCode = errorCode;
        Param = param;
    }

    public ResultError(ErrorCode errorCode, string param, string param2)
    {
        DateTimeCreation = DateTime.UtcNow;
        ErrorCode = errorCode;
        Param = param;
        Param2 = param2;
    }

    public ResultError(ErrorCode errorCode, Exception exception, string param)
    {
        DateTimeCreation = DateTime.UtcNow;
        ErrorCode = errorCode;
        Param = param;
    }

    public ResultError(ErrorCode errorCode, int lineNum, int colNum, string param)
    {
        DateTimeCreation = DateTime.UtcNow;
        ErrorCode = errorCode;
        Param = param;
        LineNum = lineNum;
        ColNum = colNum;
    }

    public ResultError(ErrorCode errorCode, int lineNum, int colNum, string param, string param2)
    {
        DateTimeCreation = DateTime.UtcNow;
        ErrorCode = errorCode;
        Param = param;
        Param2 = param2;
        LineNum = lineNum;
        ColNum = colNum;
    }

    public ResultError(ErrorCode errorCode, int lineNum, int colNum, Exception exception, string param)
    {
        DateTimeCreation = DateTime.UtcNow;
        ErrorCode = errorCode;
        Exception = exception;
        Param = param;
        LineNum = lineNum;
        ColNum = colNum;
    }

    public ResultError(ErrorCode errorCode, string fileName, int sheetNum, int colNum, CellRawValueType cellValueType)
    {
        if (string.IsNullOrWhiteSpace(fileName)) fileName = string.Empty;
        ErrorCode = errorCode;
        FileName = fileName;
        SheetNum = sheetNum;
        ColNum = colNum;
        CellValueType = cellValueType;
    }

    /// <summary>
    /// When the error was created.
    /// </summary>
    public DateTime DateTimeCreation { get; private set; }

    /// <summary>
    /// The code of the error.
    /// </summary>
    public ErrorCode ErrorCode { get; set; }

    public string FileName { get; set; } = string.Empty;

    public int SheetNum { get; set; }

    public int LineNum { get; set; } = 0;
    public int ColNum { get; set; } = 0;

    public CellRawValueType CellValueType { get; set; } = CellRawValueType.Unknow;

    public Exception? Exception { get; set; } = null;

    /// <summary>
    /// Parameter 1 of the error.
    /// </summary>
    public string Param { get; set; } = string.Empty;

    /// <summary>
    /// Parameter 2 of the error.
    /// </summary>
    public string Param2 { get; set; } = string.Empty;

    public int Counter { get; set; } = 1;

    public int IncCounter()
    {
        Counter++;
        return Counter;
    }

}