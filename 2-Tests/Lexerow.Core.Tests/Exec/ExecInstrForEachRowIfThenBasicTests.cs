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
    /// If D.Cell>50 Then D.Cell=12
    /// </summary>
    [TestMethod]
    public void Greater50Set12Ok()
    {
        LexerowCore core = new LexerowCore();

        string fileName = @"10-Files\Greater50Set12.xlsx";

        ExecResult execResult = core.Builder.CreateInstrOpenExcel("file", fileName);

        //--Create: col D, D.Cell > 50 
        InstrCompCellVal instrCompIf = core.Builder.CreateInstrCompCellVal(3, InstrCompCellValOperator.GreaterThan, 50);

        //--Create: colD, D.Cell= 12
        InstrSetCellVal instrSetValThen = core.Builder.CreateInstrSetCellVal(3, 12);
        List<InstrBase> listInstrThen = [instrSetValThen];

        // If D.Cell>50 Then D.Cell= 12
        InstrIfColThen instrIfColThen;
        execResult = core.Builder.CreateInstrIfColThen(instrCompIf, instrSetValThen, out instrIfColThen);
        Assert.IsTrue(execResult.Result);

        // ForEeach Row IfColThen
        execResult = core.Builder.CreateInstrForEachRowIfThen("file", 0, 1, instrIfColThen);
        Assert.IsTrue(execResult.Result);

        //core.Exec.FireEvent = EventOccured;

        execResult = core.Exec.Compile();
        Assert.IsTrue(execResult.Result);

        execResult = core.Exec.Execute();
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

        ExecResult execResult = core.Builder.CreateInstrOpenExcel("file", fileName);

        //--Create Comp instr: B.Cell=Null
        InstrCompCellValIsNull instrCompIf = core.Builder.CreateInstrCompCellValIsNull(1, InstrCompCellValOperator.Equal);

        //--Create: B.Cell="NA"
        InstrSetCellVal instrSetValThen = core.Builder.CreateInstrSetCellVal(1, "NA");

        // If B.Cell= null Then B.Cell= "NA"
        InstrIfColThen instrIfColThen;
        execResult = core.Builder.CreateInstrIfColThen(instrCompIf, instrSetValThen, out instrIfColThen);
        Assert.IsTrue(execResult.Result);

        // ForEeach Row IfColThen
        execResult = core.Builder.CreateInstrForEachRowIfThen("file", 0, 1, instrIfColThen);
        Assert.IsTrue(execResult.Result);

        //core.Exec.FireEvent = EventOccured;

        execResult = core.Exec.Compile();
        Assert.IsTrue(execResult.Result);

        execResult = core.Exec.Execute();
        Assert.IsTrue(execResult.Result);

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

        ExecResult execResult = core.Builder.CreateInstrOpenExcel("file", fileName);

        //--Create Comp instr: B.Cell=Null
        InstrCompCellValIsNull instrCompIf = core.Builder.CreateInstrCompCellValIsNull(1, InstrCompCellValOperator.Equal);

        //--Create: C.Cell= "09/02/2025 00:00:00"
        InstrSetCellVal instrSetValThen = core.Builder.CreateInstrSetCellVal(2, new DateTime(2025,02,09));

        // If B.Cell= null Then C.Cell= "09/02/2025 00:00:00"
        InstrIfColThen instrIfColThen;
        execResult = core.Builder.CreateInstrIfColThen(instrCompIf, instrSetValThen, out instrIfColThen);
        Assert.IsTrue(execResult.Result);

        // ForEeach Row IfColThen
        execResult = core.Builder.CreateInstrForEachRowIfThen("file", 0, 1, instrIfColThen);
        Assert.IsTrue(execResult.Result);

        execResult = core.Exec.Compile();
        Assert.IsTrue(execResult.Result);

        execResult = core.Exec.Execute();
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
