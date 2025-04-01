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
// In col(D) if cell.Value > 50 Then Cell.Value=12
//
// file=OpenExcel(fileName)
// OnExcelForEachRowIfThen(file, 0, D, >50, =12)
//
void Test1()
{
    LexerowCore core = new LexerowCore();
    core.AppTraceEvent = AppTraceEvent;

    //--Instr: file=OpenExcel(fileName)
    string fileName = @"Input\BasicDataTable.xlsx";
    ExecResult execRes = core.Builder.CreateInstrOpenExcel("file", fileName);
    if(!execRes.Result)
    {
        Console.WriteLine("=>Error occured!");
        return;
    }

    //--Create: col D, c.Value > 50 
    InstrCompColCellVal instrCompIf = core.Builder.CreateInstrCompCellVal(3, InstrCompValOperator.GreaterThan, 50);

    //--Create: colD, c.Value:=12
    InstrSetCellVal instrSetValThen = core.Builder.CreateInstrSetCellVal(3, 12);

    // If A.Cell < 10 Then B.Cell= 10
    InstrIfColThen instrIfColThen;
    execRes = core.Builder.CreateInstrIfColThen(instrCompIf, instrSetValThen, out instrIfColThen);

    execRes = core.Builder.CreateInstrOnExcelForEachRowIfThen("file", 0, 1, instrIfColThen);
    if(!execRes.Result)
    {
        Console.WriteLine("ERR, Unable to create the instrForIfThen!");
        return;
    }

    //core.Exec.EventOccurs = EventOccured;

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

// 	si est dans le range de valeur
//If InRange(c.Value, valMin, valMax)


// InstrInRangeCellVal
// 
void TestInRange()
{

}


void Test2()
{
    LexerowCore core = new LexerowCore();

    string fileName = @"Input\Test2.xlsx";
   // ExecResultOpenExcel res = core.Exec.OpenExcel(fileName);

    //if (res.Error != null)
    //{
    //    Console.WriteLine("=>Error occured!");
    //    return;
    //}

    //core.Exec.EventOccurs = EventOccured;

    // by default, the header is on row 1 

    // In col(D) if cell.Value > 15 Then Cell.Value=14
    InstrCompColCellVal exprComp = core.Builder.CreateInstrCompCellVal(3, InstrCompValOperator.GreaterThan, 15);
    //var res2 = core.Exec.SetCellsIf(res.ExcelFile, 0, "D", exprComp, 14);

    //if (res.Error == null)
    //    Console.WriteLine("Finished with success.");
    //else Console.WriteLine("Finished with error!");

}

void TestDateTime()
{
    DateTime dt;
    DateOnly d;
    // Hour only??

    //--DateOnly
    // In col(D) if cell.Value > 01/01/2024 Then Cell.Value=31/12/2023
    //ici();

}

void TestItemList()
{
    LexerowCore core = new LexerowCore();

    string fileName = @"Input\Test2.xlsx";
    //ExecResultOpenExcel res = core.Exec.OpenExcel(fileName);

    //if (res.Error != null)
    //{
    //    Console.WriteLine("=>Error occured!");
    //    return;
    //}

    // If cell.Value in [ "yes", "y", "ok" ] Then SetCell Value= "X"
    List<string> listYes = ["yes", "y", "ok"];
    //var res2 = core.Exec.SetCellsIf(res.ExcelFile, 0, "A", listYes, false, "X");

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

Test1();

//Test2();

Console.WriteLine("ends.");

