using Lexerow.Core.System;
using Lexerow.Core.Tests._20_Utils;
using NPOI.SS.Formula.Functions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Tests.Exec;

// ExecInstrForEachIfThenBasicTests
[TestClass]
public class ExecInstrForEachRowIfThenBasicTests
{
    /// <summary>
    /// file=OpenExcel("10-Files\Greater50Set12.xlsx")
    /// 
    /// OnExcel file OnSheet 0,1
    ///   ForEach Row
    ///      If D.Cell>50 Then D.Cell=12
    ///      
    /// 
    /// </summary>
    //[TestMethod]
    //public void Greater50Set12Ok()
    //{
    //    LexerowCore core = new LexerowCore();

    //    string fileName = @"10-Files\Greater50Set12.xlsx";

    //    //--instr: file=OpenExcel("10-Files\Greater50Set12.xlsx")
    //    ExecResult execResult = core.ProgBuilder.CreateInstrOpenExcelParamConst("file", fileName);

    //    // OnExcel file OnSheet 0,1	
    //    execResult = core.ProgBuilder.CreateInstrOnExcelOnSheet("file", 0, 1, out BuildInstrOnExcelOnSheet buildInstrOnExcel);

    //    //--instr: If D.Cell>50 Then D.Cell=12

    //    // TODO: rework params!!
    //    execResult = core.ProgBuilder.CreateInstrIf(buildInstrOnExcel, "D", "Cell", ">", 50, out BuildInstrOnExcelIf buildInstrIf);

    //    execResult = core.ProgBuilder.CreateInstrThenSetVar(buildInstrOnExcel, buildInstrIf, "D.Cell", 12);

    //    // save the instr OnExcel OnSheet ForEach Row IfThen
    //    execResult = core.ProgBuilder.SaveInstrOnExcelOnSheet(buildInstrOnExcel);

    //    //-execute the program
    //    execResult = core.ExecuteProgram();
    //    Assert.IsTrue(execResult.Result);

    //    // check the content of modified excel file
    //    var stream = ExcelChecker.OpenExcel(fileName);
    //    var wb = ExcelChecker.GetWorkbook(stream);
    //    bool res = ExcelChecker.CheckCellValue(wb, 0, 2, 3, 12);
    //    Assert.IsTrue(res);

    //    ExcelChecker.CloseExcel(stream);
    //}

    //--test with an undefined file varname
    // will be detected during the compilation stage

    //--test with the filenname, not var
    // OnExcel "10-Files\Greater50Set12.xlsx"



    /// <summary>
    /// If D.Cell>50 Then D.Cell=12
    /// </summary>
    [TestMethod]
    public void Greater50Set12Ok_OLD()
    {
        LexerowCore core = new LexerowCore();

        string fileName = @"10-Files\Greater50Set12.xlsx";

        ExecResult execResult = core.ProgBuilder.CreateInstrOpenExcelParamConst("file", fileName);

        //--Comparison: D.Cell > 50 
        InstrCompColCellVal instrCompIf = core.ProgBuilder.CreateInstrCompCellVal(3, ValCompOperator.GreaterThan, 50);

        //--Set: D.Cell= 12
        InstrSetCellVal instrSetValThen = core.ProgBuilder.CreateInstrSetCellVal(3, 12);
        List<InstrBase> listInstrThen = [instrSetValThen];

        // If D.Cell>50 Then D.Cell= 12
        InstrIfColThen instrIfColThen;
        execResult = core.ProgBuilder.CreateInstrIfColThen(instrCompIf, instrSetValThen, out instrIfColThen);
        Assert.IsTrue(execResult.Result);

        // ForEeach Row IfColThen
        execResult = core.ProgBuilder.CreateInstrOnExcelForEachRowIfThen("file", 0, 1, instrIfColThen);
        Assert.IsTrue(execResult.Result);

        //core.Exec.FireEvent = EventOccured;

        execResult = core.ExecuteProgram();
        Assert.IsTrue(execResult.Result);

        // check the content of modified excel file
        var stream = ExcelChecker.OpenExcel(fileName);
        var wb= ExcelChecker.GetWorkbook(stream);
        bool res= ExcelChecker.CheckCellValue(wb, 0, 2, 3, 12);

        Assert.IsTrue(res);

        ExcelChecker.CloseExcel(stream);
    }

    /// <summary>
    /// If B.Cell=Null Then B.Cell= "NA"
    /// 0   value is null?
    /// 1   azerty
    /// 2            -> null, =>NA
    /// 3   querty
    /// 4           -> empty, string, =>NA
    /// 5           -> empty, number, =>NA pb style!
    /// </summary>
    [TestMethod]
    public void ForIFCellIsNullSetNA()
    {
        LexerowCore core = new LexerowCore();

        string fileName = @"10-Files\ForIfCellIsNullSetNA.xlsx";

        ExecResult execResult = core.ProgBuilder.CreateInstrOpenExcelParamConst("file", fileName);

        //--Comp: B.Cell=Null
        InstrCompColCellValIsNull instrCompIf = core.ProgBuilder.CreateInstrCompCellValIsNull(1);

        //--Set: B.Cell="NA"
        InstrSetCellVal instrSetValThen = core.ProgBuilder.CreateInstrSetCellVal(1, "NA");

        // If B.Cell= null Then B.Cell= "NA"
        InstrIfColThen instrIfColThen;
        execResult = core.ProgBuilder.CreateInstrIfColThen(instrCompIf, instrSetValThen, out instrIfColThen);
        Assert.IsTrue(execResult.Result);

        // ForEeach Row IfColThen
        execResult = core.ProgBuilder.CreateInstrOnExcelForEachRowIfThen("file", 0, 1, instrIfColThen);
        Assert.IsTrue(execResult.Result);

        //core.Exec.FireEvent = EventOccured;

        execResult = core.ExecuteProgram();
        Assert.IsTrue(execResult.Result);

        // check the number of if condition which are matched
        Assert.AreEqual(5, execResult.Insights.AnalyzedDatarowCount);
        Assert.AreEqual(3, execResult.Insights.IfCondMatchCount);

        //--check the content of the modified excel file
        var stream = ExcelChecker.OpenExcel(fileName);
        var wb = ExcelChecker.GetWorkbook(stream);

        // row2, col1
        bool res = ExcelChecker.CheckCellValue(wb, 0, 2, 1, "NA");
        Assert.IsTrue(res);

        // row4, col1
        res = ExcelChecker.CheckCellValue(wb, 0, 4, 1, "NA");
        Assert.IsTrue(res);

        // row5, col1
        res = ExcelChecker.CheckCellValue(wb, 0, 5, 1, "NA");
        Assert.IsTrue(res);

        ExcelChecker.CloseExcel(stream);
    }

    /// <summary>
    /// ForEachRowCell If B.Cell=null Then C.Cell="09/02/2025"  
    /// 
    /// but cell contains: string, double, currency,...
    /// </summary>
    [TestMethod]
    public void ForEachRowIfThenSetDate()
    {
        LexerowCore core = new LexerowCore();

        string fileName = @"10-Files\ForIfCellIsNullSetDate.xlsx";

        ExecResult execResult = core.ProgBuilder.CreateInstrOpenExcelParamConst("file", fileName);

        //--Create Comp instr: B.Cell=Null
        InstrCompColCellValIsNull instrCompIf = core.ProgBuilder.CreateInstrCompCellValIsNull(1);

        //--Create: C.Cell= "09/02/2025 00:00:00"
        InstrSetCellVal instrSetValThen = core.ProgBuilder.CreateInstrSetCellVal(2, new DateTime(2025,02,09));

        // If B.Cell= null Then C.Cell= "09/02/2025 00:00:00"
        InstrIfColThen instrIfColThen;
        execResult = core.ProgBuilder.CreateInstrIfColThen(instrCompIf, instrSetValThen, out instrIfColThen);
        Assert.IsTrue(execResult.Result);

        // ForEeach Row IfColThen
        execResult = core.ProgBuilder.CreateInstrOnExcelForEachRowIfThen("file", 0, 1, instrIfColThen);
        Assert.IsTrue(execResult.Result);

        execResult = core.ExecuteProgram();
        Assert.IsTrue(execResult.Result);

        //--start checks
        //--check the content of the modified excel file
        var stream = ExcelChecker.OpenExcel(fileName);
        var wb = ExcelChecker.GetWorkbook(stream);

        // C3: row2, col2 ="09/02/2025 00:00:00" format= 22= "m/d/yy h:mm"
        bool res = ExcelChecker.CheckCellValue(wb, 0, 2, 2, new DateTime(2025, 02, 09));
        Assert.IsTrue(res);

        // C4: row3, col2 ="09/02/2025 00:00:00"
        res = ExcelChecker.CheckCellValue(wb, 0, 3, 2, new DateTime(2025, 02, 09));
        Assert.IsTrue(res);

        ExcelChecker.CloseExcel(stream);
    }
    // Apply 2 ForIf instructions!!

}
