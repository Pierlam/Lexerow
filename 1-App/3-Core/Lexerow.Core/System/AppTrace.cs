using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;

public enum AppTraceLevel
{
    Error,
    Warning,
    Info,
    Detail
}

public class AppTrace
{
    public AppTrace(AppTraceLevel level, string msg)
    {
        Level = level;
        Msg = msg;
    }

    public DateTime When { get; private set; } = DateTime.Now;

    public AppTraceLevel Level { get; set; } = AppTraceLevel.Info;

    // ModuleName?

    public string Msg { get; set; } = string.Empty;


}
