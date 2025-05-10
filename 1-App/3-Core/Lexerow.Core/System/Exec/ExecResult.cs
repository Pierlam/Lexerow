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

    public void AddError(ExecResultError error)
    { 
        this.ListError.Add(error);
        Result = false;
    }

    public void AddWarning(ExecResultWarning warning)
    {
        this.ListWarning.Add(warning);
        Result = false;
    }

    public void AddListError(List<ExecResultError> listError)
    {
        if (listError.Count == 0) return;

        this.ListError.AddRange(listError);
        Result = false;
    }

}
