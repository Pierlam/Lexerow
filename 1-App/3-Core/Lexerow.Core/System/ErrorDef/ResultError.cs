using OpenExcelSdk;

namespace Lexerow.Core.System;

/// <summary>
/// Execution (load, save, compilation or execution) result error.
/// </summary>
public class ResultError
{
    public ResultError(ErrorCode errorCode, string param)
    {
        ErrorCode = errorCode;
        Param = param;
    }

    public ResultError(ErrorCode errorCode, int lineNum, int colNum, string param)
    {
        ErrorCode = errorCode;
        LineNum = lineNum;
        ColNum = colNum;
        Param = param;
    }

    public ResultError(ErrorCode errorCode, int lineNum, int colNum, string param, string param2)
    {
        ErrorCode = errorCode;
        LineNum = lineNum;
        ColNum = colNum;
        Param = param;
        Param2 = param2;
    }

    public ResultError(ErrorCode errorCode, int lineNum, int colNum, Exception exception, string param)
    {
        ErrorCode = errorCode;
        LineNum = lineNum;
        ColNum = colNum;
        Exception = exception;
        Param = param;
    }

    /// <summary>
    /// TODO: pose pb!
    /// </summary>
    /// <param name="errorCode"></param>
    /// <param name="fileName"></param>
    /// <param name="sheetNum"></param>
    /// <param name="colNum"></param>
    /// <param name="cellValueType"></param>
    public ResultError(ErrorCode errorCode, string fileName, int sheetNum, int colNum, ExcelCellType cellValueType)
    {
        if (string.IsNullOrWhiteSpace(fileName)) fileName = string.Empty;
        ErrorCode = errorCode;
        //LineNum = lineNum;
        ColNum = colNum;
        FileName = fileName;
        SheetNum = sheetNum;
        CellValueType = cellValueType;
    }

    /// <summary>
    /// When the error was created.
    /// </summary>
    public DateTime DateTimeCreation { get; private set; } = DateTime.UtcNow;

    /// <summary>
    /// The code of the error.
    /// </summary>
    public ErrorCode ErrorCode { get; set; }

    public string FileName { get; set; } = string.Empty;

    public int SheetNum { get; set; }

    public int LineNum { get; set; } = 0;
    public int ColNum { get; set; } = 0;

    public ExcelCellType CellValueType { get; set; } = ExcelCellType.Undefined;

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