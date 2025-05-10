using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;

/// <summary>
///  Execution  result warning. 
///  exp: IfCondTypeMismatch, excelFile=file, SheetNum=0, col=A, type=Number, Counter=12
/// </summary>
public class ExecResultWarning
{
    public ExecResultWarning(ErrorCode errorCode, string param)
    {
        DateTimeCreation = DateTime.UtcNow;
        ErrorCode = errorCode;
        Param = param;
    }

    public ExecResultWarning(ErrorCode errorCode)
    {
        DateTimeCreation = DateTime.UtcNow;
        ErrorCode = errorCode;
    }

    /// <summary>
    /// When the error was created.
    /// </summary>
    public DateTime DateTimeCreation { get; private set; }

    /// <summary>
    /// Type of the error: BuildInstr, ExecInstr, CompileScript
    /// and more: License config.
    /// </summary>
    //public ErrorType ErrorType { get; set; }

    /// <summary>
    /// The code of the warning.
    /// </summary>
    public ErrorCode ErrorCode { get; set; }

    /// <summary>
    /// Parameter 1 of the error.
    /// </summary>
    public string Param { get; set; } = string.Empty;

    /// <summary>
    /// Parameter 2 of the error.
    /// </summary>
    public string Param2 { get; set; } = string.Empty;

}
