namespace Lexerow.Core.System;

/// <summary>
///  Result warning.
///  exp: IfCondTypeMismatch, excelFile=file, SheetNum=0, col=A, type=Number, Counter=12
/// </summary>
public class ResultWarning
{
    public ResultWarning(ErrorCode errorCode, string param)
    {
        ErrorCode = errorCode;
        Param = param;
    }

    public ResultWarning(ErrorCode errorCode)
    {
        ErrorCode = errorCode;
    }

    public ResultWarning(ErrorCode errorCode, string fileName, int sheetNum, int colNum, CellRawValueType cellValueType)
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
    public DateTime DateTimeCreation { get; private set; } = DateTime.UtcNow;

    /// <summary>
    /// Type of the error: BuildInstr, ExecInstr, CompileScript
    /// and more: License config.
    /// </summary>
    //public ErrorType ErrorType { get; set; }

    /// <summary>
    /// The code of the warning.
    /// </summary>
    public ErrorCode ErrorCode { get; set; }

    public string FileName { get; set; } = string.Empty;

    public int SheetNum { get; set; }

    public int ColNum { get; set; }

    public CellRawValueType CellValueType { get; set; } = CellRawValueType.Unknow;

    public int Counter { get; set; } = 1;

    /// <summary>
    /// Parameter 1 of the error.
    /// </summary>
    public string Param { get; set; } = string.Empty;

    /// <summary>
    /// Parameter 2 of the error.
    /// </summary>
    public string Param2 { get; set; } = string.Empty;

    public int IncCounter()
    {
        Counter++;
        return Counter;
    }
}