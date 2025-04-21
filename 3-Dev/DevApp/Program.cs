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
void Test1()
{
    LexerowCore core = new LexerowCore();
    core.AppTraceEvent = AppTraceEvent;

    //--Create: file=OpenExcel(fileName)
    string fileName = @"Input\BasicDataTable.xlsx";
    ExecResult execRes = core.Builder.CreateInstrOpenExcel("file", fileName);
    if(!execRes.Result)
    {
        Console.WriteLine("=>Error occured!");
        return;
    }

    //--Comp: D.Cell > 50 
    InstrCompColCellVal instrCompIf = core.Builder.CreateInstrCompCellVal(3, ValCompOperator.GreaterThan, 50);

    //--Set: D.Cell= 12
    InstrSetCellVal instrSetValThen = core.Builder.CreateInstrSetCellVal(3, 12);

    //--If D.Cell > 50 Then D.Cell= 12
    InstrIfColThen instrIfColThen;
    execRes = core.Builder.CreateInstrIfColThen(instrCompIf, instrSetValThen, out instrIfColThen);

    //--Create: OnExcel ForEach Row
    execRes = core.Builder.CreateInstrOnExcelForEachRowIfThen("file", 0, 1, instrIfColThen);
    if(!execRes.Result)
    {
        Console.WriteLine("ERR, Unable to create the instrForIfThen!");
        return;
    }


    execRes = core.Exec.Compile();
    if (!execRes.Result)
    {
        Console.WriteLine("ERR, Unable to compile the source code");
        return;
    }

    //--Execute all saved instruction
    execRes = core.Exec.Execute();

    if (execRes.Result)
        Console.WriteLine("Finished with success.");
    else Console.WriteLine("Finished with error!");

}


///
/// If A.Cell In [ "y", "yes", "ok" ] Then A.Cell= "X"
/// by default, it is case sensitive.
/// 
/// Case insensitive:
/// If A.Cell In /I [ "y", "yes", "ok" ] Then A.Cell= "X"
///  In /CI -> In case Insensitive
///  In    -> In case Sensitive
void TestFuncInListOfItems()
{
    LexerowCore core = new LexerowCore();

    //--Create: file=OpenExcel(fileName)
    string fileName = @"Input\TestFuncInListOfItems.xlsx";
    ExecResult execRes = core.Builder.CreateInstrOpenExcel("file", fileName);
    if (!execRes.Result)
    {
        Console.WriteLine("=>Error occured!");
        return;
    }

    List<string> listYes = ["yes", "y", "ok"];

    //--Comp: A.Cell In [ "y", "yes", "ok" ]
    //InstrCompColCellVal instrCompIf = core.Builder.CreateInstrCompCellVal(0, ValCompOperator.GreaterThan, 50);
    // A.Cell, listOfItems, In=true/NotIn=false, case sensitive=true
    execRes = core.Builder.CreateInstrCompCellInItems(0, listYes, true, out InstrCompColCellInStringItems instrCompIf);

    //--Set: A.Cell= "X"
    InstrSetCellVal instrSetValThen = core.Builder.CreateInstrSetCellVal(0, "X");

    //--If A.Cell In [ "y", "yes", "ok" ] Then A.Cell= "X"
    InstrIfColThen instrIfColThen;
    execRes = core.Builder.CreateInstrIfColThen(instrCompIf, instrSetValThen, out instrIfColThen);

    //--Create: OnExcel ForEach Row
    execRes = core.Builder.CreateInstrOnExcelForEachRowIfThen("file", 0, 1, instrIfColThen);
    if (!execRes.Result)
    {
        Console.WriteLine("ERR, Unable to create the instrForIfThen!");
        return;
    }


    execRes = core.Exec.Compile();
    if (!execRes.Result)
    {
        Console.WriteLine("ERR, Unable to compile the source code");
        return;
    }

    //--Execute all saved instruction
    execRes = core.Exec.Execute();

    if (execRes.Result)
        Console.WriteLine("Finished with success.");
    else Console.WriteLine("Finished with error!");


}

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

TestFuncInListOfItems();

Console.WriteLine("ends.");

