using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System.Exec.Event;
public class InstrOpenExcelExecEvent : InstrBaseExecEvent
{
    public InstrOpenExcelExecEvent()
    {
        Type = InstrBaseExecEventType.OpenExcel;
    }

    public string FileName { get; set; }

    public static InstrOpenExcelExecEvent CreateStart(string fileName)
    {
        var execEvent= new InstrOpenExcelExecEvent();
        execEvent.State = InstrBaseExecEventState.Start;
        execEvent.Result = InstrBaseExecEventResult.Ok;
        execEvent.FileName = fileName;
        return execEvent;
    }

    public static InstrOpenExcelExecEvent CreateFinished(DateTime dtStart, InstrBaseExecEventResult result, string fileName)
    {
        var execEvent = new InstrOpenExcelExecEvent();
        execEvent.State = InstrBaseExecEventState.Finished;
        execEvent.Result = result;
        execEvent.FileName = fileName;
        execEvent.ElapsedTime= DateTime.Now - dtStart;
        return execEvent;
    }

}
