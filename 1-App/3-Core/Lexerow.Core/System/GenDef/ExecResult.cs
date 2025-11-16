using Lexerow.Core.System.ScriptDef;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;

/// <summary>
/// Execution result.
/// Function/process/treatment execution result.
/// 
/// </summary>
public class ExecResult
{
    public ExecResult()
    {
    }

    /// <summary>
    /// Result of the execution. 
    /// Return true if there is no error.
    /// true by default.
    /// </summary>
    public bool Result { get; set; } = true;

    /// <summary>
    /// list of error occured.
    /// </summary>
    public List<ExecResultError> ListError { get; private set; } = new List<ExecResultError>();

    /// <summary>
    /// list of warning occured.
    /// </summary>
    public List<ExecResultWarning> ListWarning { get; private set; } = new List<ExecResultWarning>();

    /// <summary>
    /// execution result insights/informations: how many datarow are modified, created or removed.
    /// </summary>
    public ExecResultInsights Insights { get; private set; } = new ExecResultInsights();


    public void AddListError(List<ExecResultError> listError)
    {
        if (listError.Count == 0) return;

        ListError.AddRange(listError);
        Result = false;
    }

    public void AddError(ExecResultError error)
    {
        this.ListError.Add(error);
        Result = false;
    }

    public ExecResultError AddError(ErrorCode errorCode, Exception exception, string msg)
    {
        var execResultError = new ExecResultError(errorCode, 0, 0, exception, msg);
        ListError.Add(execResultError);
        Result = false;
        return execResultError;
    }

    public ExecResultError AddError(ErrorCode errorCode, string msg)
    {
        var execResultError = new ExecResultError(errorCode, 0, 0, msg);
        ListError.Add(execResultError);
        Result = false;
        return execResultError;
    }

    public ExecResultError AddError(ErrorCode errorCode, int numLine, int colLine, string msg)
    {
        var execResultError = new ExecResultError(errorCode, numLine, 0, msg);
        ListError.Add(execResultError);
        Result = false;
        return execResultError;
    }

    /// <summary>
    /// Error occurs during script compilation.
    /// </summary>
    /// <param name="errorCode"></param>
    /// <param name="scriptToken"></param>
    public void AddError(ErrorCode errorCode, ScriptToken scriptToken)
    {
        ExecResultError execResultError;
        if (scriptToken != null)
            execResultError = new ExecResultError(errorCode, scriptToken.LineNum, scriptToken.ColNum, scriptToken.Value);
        else
            execResultError = new ExecResultError(errorCode, 0, 0, string.Empty);
        ListError.Add(execResultError);
        Result = false;
    }

    /// <summary>
    /// Error occurs during script compilation.
    /// </summary>
    /// <param name="errorCode"></param>
    /// <param name="scriptToken"></param>
    public void AddError(ErrorCode errorCode, ScriptToken scriptToken, string param)
    {
        ExecResultError execResultError;
        if (scriptToken!=null)
            execResultError = new ExecResultError(errorCode, scriptToken.LineNum, scriptToken.ColNum, scriptToken.Value, param);
        else
            execResultError = new ExecResultError(errorCode, 0, 0, string.Empty, param);
        ListError.Add(execResultError);
        Result = false;
    }

    /// <summary>
    /// Error occurs during script compilation.
    /// </summary>
    /// <param name="errorCode"></param>
    /// <param name="scriptToken"></param>
    public void AddError(ErrorCode errorCode, ScriptToken scriptToken, Exception exception)
    {
        ExecResultError execResultError;
        if (scriptToken != null)
            execResultError = new ExecResultError(errorCode, scriptToken.LineNum, scriptToken.ColNum, exception, scriptToken.Value);
        else
            execResultError = new ExecResultError(errorCode, 0, 0, exception, string.Empty);
        ListError.Add(execResultError);
        Result = false;
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
    public ExecResultWarning? FindWarning(ErrorCode errorCode, string fileName, int sheetNum, int colNum, CellRawValueType cellValueType)
    {
        if(fileName== null)fileName = string.Empty; 
        return ListWarning.Find(x => x.ErrorCode ==errorCode && x.FileName.Equals(fileName,StringComparison.CurrentCultureIgnoreCase) && x.SheetNum== sheetNum && x.ColNum== colNum && x.CellValueType== cellValueType);
    }

    public void AddWarning(ErrorCode errorCode, string fileName, int sheetNum, int colNum, CellRawValueType cellValueType)
    {
        if (string.IsNullOrWhiteSpace(fileName)) fileName = string.Empty;

        // is there a warning already existing? Code=IfCondTypeMismatch, ExcelFile, fileName, SheetNum, colNum, valType
        ExecResultWarning? warning = FindWarning(errorCode, fileName, sheetNum, colNum, cellValueType);
        if (warning != null)
        {
            warning.IncCounter();
            return;
        }

        warning = new ExecResultWarning(errorCode, fileName, sheetNum, colNum, cellValueType);
        ListWarning.Add(warning);
    }
}
