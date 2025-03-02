using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System.Exec.Event;


public enum InstrBaseExecEventType
{
    OpenExcel,
    CloseExcel,
    ForEachRowCellIfThen
}

public enum InstrBaseExecEventState
{
    Start,
    InProgress,
    Finished
}

public enum InstrBaseExecEventResult
{
    Ok,
    Error
}

public abstract class InstrBaseExecEvent
{
    public InstrBaseExecEventType Type { get; set; }

    public InstrBaseExecEventState State { get; set; }= InstrBaseExecEventState.Start;

    public InstrBaseExecEventResult Result { get; set; } = InstrBaseExecEventResult.Ok;

    public DateTime When { get; set; }= DateTime.Now;

    public TimeSpan ElapsedTime { get; set; }= TimeSpan.Zero;

}
