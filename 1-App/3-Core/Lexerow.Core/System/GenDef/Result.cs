using Lexerow.Core.System.ScriptDef;

namespace Lexerow.Core.System;

/// <summary>
/// Result of a function or a program execution.
///
/// </summary>
public class Result
{
    public Result()
    {
        
    }

    /// <summary>
    /// Result of the execution.
    /// Return true if there is no error.
    /// true by default.
    /// </summary>
    public bool Res { get; set; } = true;

    /// <summary>
    /// list of error occured.
    /// </summary>
    public List<ResultError> ListError { get; private set; } = new List<ResultError>();

    /// <summary>
    /// list of warning occured.
    /// </summary>
    public List<ResultError> ListWarning { get; private set; } = new List<ResultError>();

    /// <summary>
    /// execution result insights/informations: how many datarow are modified, created or removed.
    /// </summary>
    public ResultInsights Insights { get; private set; } = new ResultInsights();

    public void AddListError(List<ResultError> listError)
    {
        if (listError.Count == 0) return;

        ListError.AddRange(listError);
        Res = false;
    }

    public void AddError(ResultError error)
    {
        this.ListError.Add(error);
        Res = false;
    }

    public ResultError AddError(ErrorCode errorCode, Exception exception, string msg)
    {
        var resultError = new ResultError(errorCode, 0, 0, exception, msg);
        ListError.Add(resultError);
        Res = false;
        return resultError;
    }

    public ResultError AddError(ErrorCode errorCode, string msg)
    {
        var resultError = new ResultError(errorCode, 0, 0, msg);
        ListError.Add(resultError);
        Res = false;
        return resultError;
    }

    public ResultError AddError(ErrorCode errorCode, int numLine, int colLine, string msg)
    {
        var resultError = new ResultError(errorCode, numLine, 0, msg);
        ListError.Add(resultError);
        Res = false;
        return resultError;
    }

    /// <summary>
    /// Error occurs during script compilation.
    /// </summary>
    /// <param name="errorCode"></param>
    /// <param name="scriptToken"></param>
    public void AddError(ErrorCode errorCode, ScriptToken scriptToken)
    {
        ResultError resultError;
        if (scriptToken != null)
            resultError = new ResultError(errorCode, scriptToken.LineNum, scriptToken.ColNum, scriptToken.Value);
        else
            resultError = new ResultError(errorCode, 0, 0, string.Empty);
        ListError.Add(resultError);
        Res = false;
    }

    public void AddWarning(ErrorCode errorCode, ScriptToken scriptToken)
    {
        ResultError resultError;
        if (scriptToken != null)
            resultError = new ResultError(errorCode, scriptToken.LineNum, scriptToken.ColNum, scriptToken.Value);
        else
            resultError = new ResultError(errorCode, 0, 0, string.Empty);
        ListWarning.Add(resultError);
        Res = false;
    }

    /// <summary>
    /// Error occurs during script compilation.
    /// </summary>
    /// <param name="errorCode"></param>
    /// <param name="scriptToken"></param>
    public void AddError(ErrorCode errorCode, ScriptToken scriptToken, string param)
    {
        ResultError resultError;
        if (scriptToken != null)
            resultError = new ResultError(errorCode, scriptToken.LineNum, scriptToken.ColNum, scriptToken.Value, param);
        else
            resultError = new ResultError(errorCode, 0, 0, string.Empty, param);
        ListError.Add(resultError);
        Res = false;
    }

    /// <summary>
    /// Error occurs during script compilation.
    /// </summary>
    /// <param name="errorCode"></param>
    /// <param name="scriptToken"></param>
    public void AddError(ErrorCode errorCode, ScriptToken scriptToken, Exception exception)
    {
        ResultError resultError;
        if (scriptToken != null)
            resultError = new ResultError(errorCode, scriptToken.LineNum, scriptToken.ColNum, exception, scriptToken.Value);
        else
            resultError = new ResultError(errorCode, 0, 0, exception, string.Empty);
        ListError.Add(resultError);
        Res = false;
    }

    /// <summary>
    /// Is there a warning already existing? Code=IfCondTypeMismatch, ExcelFile, fileName, SheetNum, colNum, valType
    /// </summary>
    /// <param name="errorCode"></param>
    /// <param name="fileName"></param>
    /// <param name="sheetNum"></param>
    /// <param name="colNum"></param>
    /// <param name="cellValueType"></param>
    /// <returns></returns>
    public ResultError? FindWarning(ErrorCode errorCode, string fileName, int sheetNum, int colNum, CellRawValueType cellValueType)
    {
        if (fileName == null) fileName = string.Empty;
        return ListWarning.Find(x => x.ErrorCode == errorCode && x.FileName.Equals(fileName, StringComparison.CurrentCultureIgnoreCase) && x.SheetNum == sheetNum && x.ColNum == colNum && x.CellValueType == cellValueType);
    }

    public void AddWarning(ErrorCode errorCode, string fileName, int sheetNum, int colNum, CellRawValueType cellValueType)
    {
        if (string.IsNullOrWhiteSpace(fileName)) fileName = string.Empty;

        // is there a warning already existing? Code=IfCondTypeMismatch, ExcelFile, fileName, SheetNum, colNum, valType
        ResultError? warning = FindWarning(errorCode, fileName, sheetNum, colNum, cellValueType);
        if (warning != null)
        {
            warning.IncCounter();
            return;
        }
        warning = new ResultError(errorCode, fileName, sheetNum, colNum, cellValueType);
        ListWarning.Add(warning);
    }
}