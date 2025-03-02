using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;

/// <summary>
/// Generic result execution.
/// </summary>
public class ExecResult
{
    public ExecResult()
    {
    }

    /// <summary>
    /// Result is ok by default.
    /// </summary>
    public bool Result { get; set; } = true;

    /// <summary>
    /// list of ErrorBase
    /// </summary>
    public List<CoreError> ListError { get; private set; } = new List<CoreError>();

    public void AddError(CoreError error)
    { 
        this.ListError.Add(error);
        Result = false;
    }
}
