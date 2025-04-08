using Lexerow.Core.System.Exec.Event;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;

public enum AppTraceModule
{
    Builder,
    Compile,
    Exec
}

public enum AppTraceLevel
{
    Error,
    Warning,
    Info,
    Detail
}

public class AppTrace
{
    public AppTrace(AppTraceModule module, AppTraceLevel level, string msg)
    {
        Module = module;
        Level = level;
        Msg = msg;
    }

    public AppTrace(AppTraceModule module, AppTraceLevel level, string msg, InstrBaseExecEvent instrExecEvent)
    {
        Module = module;
        Level = level;
        Msg = msg;
        InstrExecEvent = instrExecEvent;
    }

    public DateTime When { get; private set; } = DateTime.Now;

    public AppTraceModule Module { get; set; } = AppTraceModule.Builder;

    public AppTraceLevel Level { get; set; } = AppTraceLevel.Info;

    // ModuleName?

    public string Msg { get; set; } = string.Empty;

    /// <summary>
    /// Set if an instruction is executed.
    /// Module=Exec
    /// </summary>
    public InstrBaseExecEvent? InstrExecEvent { get; set; } = null;
}
