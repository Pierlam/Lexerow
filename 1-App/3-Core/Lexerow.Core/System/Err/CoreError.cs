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
        Message = msg;
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


    public string Message { get; set; }=string.Empty;
}
