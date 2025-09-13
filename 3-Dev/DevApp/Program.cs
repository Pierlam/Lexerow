using DevApp;
using Lexerow.Core;
using Lexerow.Core.System;
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


///
//
// OnExcel fileName
//   OnEach Row
//     If D.Cell > 50 Then D.Cell= 12
//
// file=OpenExcel(fileName)
// OnExcelForEachRowIfThen(file, 0, D, >50, =12)
//


///
/// If A.Cell In [ "y", "yes", "ok" ] Then A.Cell= "X"
/// by default, it is case sensitive.
/// 
/// Case insensitive:
/// If A.Cell In [ "y", "yes", "ok" ] /I Then A.Cell= "X"
///  In /CI -> In case Insensitive
///  In    -> In case Sensitive

//XXXXXXXXXXXXXXXXXXXX-MAIN:

DevNpoi devNpoi = new DevNpoi();

//string fileName = @"Input\Test1.xlsx";
//devNpoi.TestTypes(fileName);

//string fileName = @"Input\TestTypes.xlsx";
//devNpoi.TestTypes();

//string fileName = @"Input\TestDate.xlsx";
//devNpoi.TestDateTime(fileName);

//devNpoi.TestBlankNull();

//Test1();

//Test2();


Console.WriteLine("ends.");

