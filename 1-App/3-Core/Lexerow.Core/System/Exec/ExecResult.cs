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

    public void AddError(ExecResultError error)
    { 
        this.ListError.Add(error);
        Result = false;
    }

    public void AddListError(List<ExecResultError> listError)
    {
        if (listError.Count == 0) return;

        this.ListError.AddRange(listError);
        Result = false;
    }

    // is there an warning already existing? Code=IfCondTypeMismatch, ExcelFile, fileName, SheetNum, colNum, valType
    public ExecResultWarning? FindWarning(ErrorCode errorCode, string fileName, int sheetNum, int colNum, CellRawValueType cellValueType)
    {
        if(fileName== null)fileName = string.Empty; 
        return ListWarning.Find(x => x.ErrorCode ==errorCode && x.FileName.Equals(fileName,StringComparison.CurrentCultureIgnoreCase) && x.SheetNum== sheetNum && x.ColNum== colNum && x.CellValueType== cellValueType);
    }

    public void AddWarning(ErrorCode errorCode, string fileName, int sheetNum, int colNum, CellRawValueType cellValueType)
    {
        if (string.IsNullOrWhiteSpace(fileName)) fileName = string.Empty;

        // is there an warning already existing? Code=IfCondTypeMismatch, ExcelFile, fileName, SheetNum, colNum, valType
        ExecResultWarning? warning = FindWarning(errorCode, fileName, sheetNum, colNum, cellValueType);
        if (warning != null)
        {
            warning.IncCounter();
            return;
        }

        warning = new ExecResultWarning(errorCode, fileName, sheetNum, colNum, cellValueType);
        this.ListWarning.Add(warning);
    }
}
