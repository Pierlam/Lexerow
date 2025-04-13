using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;
public class CoreError
{
    public CoreError(ErrorCode errorCode, string msg)
    {
        DateTimeCreation=DateTime.UtcNow;
        ErrorCode = errorCode;
        Param = msg;
    }

    public CoreError(ErrorCode errorCode)
    {
        DateTimeCreation = DateTime.UtcNow;
        ErrorCode = errorCode;
    }

    /// <summary>
    /// When the error was created.
    /// </summary>
    public DateTime DateTimeCreation { get; private set; }

    /// <summary>
    /// Type of the error: License, parse, Exec or config.
    /// </summary>
    //public ErrorType ErrorType { get; set; }

    /// <summary>
    /// The code of the error.
    /// </summary>
    public ErrorCode ErrorCode { get; set; }

    /// <summary>
    /// Parameter 1 of the error.
    /// </summary>
    public string Param { get; set; }=string.Empty;

    /// <summary>
    /// Parameter 2 of the error.
    /// </summary>
    public string Param2 { get; set; } = string.Empty;
}
