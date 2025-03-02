using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System.Exec.Event;
public class InstrForEachRowIfThenExecEvent : InstrBaseExecEvent
{
    public InstrForEachRowIfThenExecEvent()
    {
        Type = InstrBaseExecEventType.ForEachRowCellIfThen;
    }

    public CoreError? Error { get; set; } = null;

    public int DataRowCount { get; set; } = 0;
    
    public int IfConditionFiredCount { get; set; } = 0; 

    public static InstrForEachRowIfThenExecEvent CreateStart()
    {
        var execEvent = new InstrForEachRowIfThenExecEvent();
        execEvent.State = InstrBaseExecEventState.Start;
        execEvent.Result = InstrBaseExecEventResult.Ok;
        //execEvent.FileName = fileName;
        return execEvent;
    }

    public static InstrForEachRowIfThenExecEvent CreateFinished(DateTime dtStart, InstrBaseExecEventResult result)
    {
        var execEvent = new InstrForEachRowIfThenExecEvent();
        execEvent.State = InstrBaseExecEventState.Finished;
        execEvent.Result = result;
        execEvent.ElapsedTime = DateTime.Now - dtStart;
        return execEvent;
    }

    public static InstrForEachRowIfThenExecEvent CreateFinishedOk(DateTime dtStart, int dataRowCount, int ifConditionFiredCount)
    {
        var execEvent = new InstrForEachRowIfThenExecEvent(); 
        execEvent.State = InstrBaseExecEventState.Finished;
        execEvent.Result = InstrBaseExecEventResult.Ok;
        execEvent.DataRowCount= dataRowCount;
        execEvent.IfConditionFiredCount= ifConditionFiredCount;
        execEvent.ElapsedTime = DateTime.Now - dtStart;
        return execEvent;
    }
    
    public static InstrForEachRowIfThenExecEvent CreateFinishedInProgress(DateTime dtStart, int dataRowCount, int ifConditionFiredCount)
    {
        var execEvent = new InstrForEachRowIfThenExecEvent();
        execEvent.State = InstrBaseExecEventState.InProgress;
        execEvent.Result = InstrBaseExecEventResult.Ok;
        execEvent.DataRowCount = dataRowCount;
        execEvent.IfConditionFiredCount = ifConditionFiredCount;
        execEvent.ElapsedTime = DateTime.Now - dtStart;
        return execEvent;
    }

    public static InstrForEachRowIfThenExecEvent CreateFinishedError(DateTime dtStart, CoreError error)
    {
        var execEvent = new InstrForEachRowIfThenExecEvent();
        execEvent.State = InstrBaseExecEventState.Finished;
        execEvent.Result = InstrBaseExecEventResult.Error;
        execEvent.Error = error;
        execEvent.ElapsedTime = DateTime.Now - dtStart;
        return execEvent;
    }

}
