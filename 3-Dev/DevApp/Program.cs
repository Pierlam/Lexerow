using DevApp;
using Lexerow.Core;
using Lexerow.Core.System;
using Lexerow.Core.System.ActivityLog;
using Lexerow.Core.System.Exec.Event;

Console.WriteLine("==>Lexerow Dev:");

void AppTraceEvent(AppTrace appTrace)
{
    Console.WriteLine(appTrace.When.ToString() + "|" + appTrace.Level.ToString() + "|" + appTrace.Msg);
}

void EventOccured(InstrBaseExecEvent execEvent)
{
    InstrOpenExcelExecEvent oe = execEvent as InstrOpenExcelExecEvent;
    if(oe != null)
    {
        Console.WriteLine(oe.When.ToString()+ " OpenExcel: " + oe.State+ ", Res: " + oe.Result + ", ElapsedTime: " + oe.ElapsedTime.ToString());
    }

    InstrForEachRowIfThenExecEvent fr = execEvent as InstrForEachRowIfThenExecEvent;
    if (fr != null) 
    {
        Console.WriteLine(fr.When.ToString() + " ForEachIf: " + fr.State + ", Res: " + fr.Result + ", Row count: " + fr.DataRowCount + ", IfCondFiredCount: " + fr.IfConditionFiredCount + ", ElapsedTime: " + fr.ElapsedTime.ToString());
    }
}


void Core_ActivityLogEvent(object? sender, ActivityLogBase e)
{
    Console.WriteLine("ActivityLogEvent!");
}

void TestCore()
{
    LexerowCore core = new LexerowCore();
    core.ActivityLogEvent += Core_ActivityLogEvent;
    core.LoadScriptFromFile("scriptName", "fileName");
}



///
/// If A.Cell In [ "y", "yes", "ok" ] Then A.Cell= "X"
/// by default, it is case sensitive.
/// 
/// Case insensitive:
/// If A.Cell In [ "y", "yes", "ok" ] /I Then A.Cell= "X"
///  In /CI -> In case Insensitive
///  In    -> In case Sensitive

//XXXXXXXXXXXXXXXXXXXX-MAIN:

//DevNpoi devNpoi = new DevNpoi();

//string fileName = @"Input\Test1.xlsx";
//devNpoi.TestTypes(fileName);

//string fileName = @"Input\TestTypes.xlsx";
//devNpoi.TestTypes();

//string fileName = @"Input\TestDate.xlsx";
//devNpoi.TestDateTime(fileName);

//devNpoi.TestBlankNull();

TestCore();

//Test2();


Console.WriteLine("ends.");

