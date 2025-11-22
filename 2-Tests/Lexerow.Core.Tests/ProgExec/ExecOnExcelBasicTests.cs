using Lexerow.Core.ExcelLayer;
using Lexerow.Core.ProgExec;
using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;
using Lexerow.Core.System.ScriptDef;
using Lexerow.Core.Tests._05_Common;
using Lexerow.Core.Tests._20_Utils;
using Lexerow.Core.Tests.Common;

namespace Lexerow.Core.Tests.ProgExec;

[TestClass]
public class ExecOnExcelBasicTests : BaseTests
{
    /// <summary>
    /// OnExcel "dataOnExcel1.xlsx"
	///   ForEach Row
    ///     If A.Cell >10 Then A.Cell= 10
    ///    Next
    /// End OnExcel
    ///
    ///--excel content:
    ///  age
    ///   9
    ///  13  -> 10
    ///
    /// </summary>
    [TestMethod]
    public void ExecOnExcelFilenameStringOk()
    {
        Program program = TestInstrBuilder.CreateProgram();

        //--OnExcel #1:
        //--If A.Cell >10 Then A.Cell= 10
        InstrIfThenElse instrIfThenElse = TestInstrBuilder.CreateInstrIfThen("A", 1, ">", 10, "A", 1, 10);

        //--OnExcel "dataOnExcel1.xlsx" ForEach Row IfThen Next
        InstrOnExcel instrOnExcel = TestInstrBuilder.CreateInstrOnExcelFileString(AddDblQuote(PathExcelFilesExec + "dataOnExcel1.xlsx"), instrIfThenElse);
        program.ListInstr.Add(instrOnExcel);

        //--create the program runner
        ActivityLogger logger = new ActivityLogger();
        ExcelProcessorNpoi excelProcessor = new ExcelProcessorNpoi();
        ProgramExecutor programExec = new ProgramExecutor(logger, excelProcessor);
        ExecResult execResult = new ExecResult();
        bool res = programExec.Exec(execResult, program);
        Assert.IsTrue(res);

        //--check the content of excel file
        var fileStream = TestExcelChecker.OpenExcel(PathExcelFilesExec + "dataOnExcel1.xlsx");
        Assert.IsNotNull(fileStream);
        var wb = TestExcelChecker.GetWorkbook(fileStream);

        // r1, c0: 9  -> not modified
        res = TestExcelChecker.CheckCellValue(wb, 0, 1, 0, 9);
        Assert.IsTrue(res);

        // r2, c0: 10 -> modified!
        res = TestExcelChecker.CheckCellValue(wb, 0, 2, 0, 10);
        Assert.IsTrue(res);
    }

    /// <summary>
    /// OnExcel "dataOnExcelJokerA*.xlsx"
	///   ForEach Row
    ///     If A.Cell >10 Then A.Cell= 10
    ///    Next
    /// End OnExcel
    ///
    ///--excel content:
    ///  age
    ///   9
    ///  13  -> 10
    ///
    /// </summary>
    [TestMethod]
    public void ExecOnExcelFilenameJokerStringOk()
    {
        Program program = TestInstrBuilder.CreateProgram();

        //--OnExcel #1:
        //--If A.Cell >10 Then A.Cell= 10
        InstrIfThenElse instrIfThenElse = TestInstrBuilder.CreateInstrIfThen("A", 1, ">", 10, "A", 1, 10);

        //--OnExcel "dataOnExcel1.xlsx" ForEach Row IfThen Next
        InstrOnExcel instrOnExcel = TestInstrBuilder.CreateInstrOnExcelFileString(AddDblQuote(PathExcelFilesExec + "dataOnExcelJokerA*.xlsx"), instrIfThenElse);
        program.ListInstr.Add(instrOnExcel);

        //--create the program runner
        ActivityLogger logger = new ActivityLogger();
        ExcelProcessorNpoi excelProcessor = new ExcelProcessorNpoi();
        ProgramExecutor programExec = new ProgramExecutor(logger, excelProcessor);
        ExecResult execResult = new ExecResult();
        bool res = programExec.Exec(execResult, program);
        Assert.IsTrue(res);

        //--check the content of excel file
        var fileStream = TestExcelChecker.OpenExcel(PathExcelFilesExec + "dataOnExcelJokerA.xlsx");
        Assert.IsNotNull(fileStream);
        var wb = TestExcelChecker.GetWorkbook(fileStream);

        // r1, c0: 9  -> not modified
        res = TestExcelChecker.CheckCellValue(wb, 0, 1, 0, 9);
        Assert.IsTrue(res);

        // r2, c0: 10 -> modified!
        res = TestExcelChecker.CheckCellValue(wb, 0, 2, 0, 10);
        Assert.IsTrue(res);

        //--check the content of excel file
        fileStream = TestExcelChecker.OpenExcel(PathExcelFilesExec + "dataOnExcelJokerA2.xlsx");
        Assert.IsNotNull(fileStream);
        wb = TestExcelChecker.GetWorkbook(fileStream);

        // r1, c0: 9  -> not modified
        res = TestExcelChecker.CheckCellValue(wb, 0, 1, 0, 9);
        Assert.IsTrue(res);

        // r2, c0: 10 -> modified!
        res = TestExcelChecker.CheckCellValue(wb, 0, 2, 0, 10);
        Assert.IsTrue(res);
    }

    /// <summary>
    ///
    /// file= "dataOnExcel2.xlsx"
    /// OnExcel file
    ///   ForEach Row
    ///     If A.Cell >10 Then A.Cell= 10
    ///   Next
    /// End OnExcel
    /// </summary>
    [TestMethod]
    public void ExecOpenExcelFileNameVarStringOk()
    {
        Program program = TestInstrBuilder.CreateProgram();

        //-->SetVar #1:   file= "data1.xslx"
        //    InstrLeft:  ObjectName: file
        //    InstrRight: OpenExcel, p="data.xlsx"

        // instr left
        InstrObjectName instrObjectName = TestInstrBuilder.CreateInstrObjectName("file");

        //instr right
        InstrValue instrValue = TestInstrBuilder.CreateValueString(AddDblQuote(PathExcelFilesExec + "dataOnExcel2.xlsx"));

        InstrSetVar instrSetVar = TestInstrBuilder.CreateInstrSetVar(instrObjectName, instrValue);
        program.ListInstr.Add(instrSetVar);

        //-->OnExcel #2:
        //--If A.Cell >10 Then A.Cell= 10
        InstrIfThenElse instrIfThenElse = TestInstrBuilder.CreateInstrIfThen("A", 1, ">", 10, "A", 1, 10);

        //--OnExcel file ForEach Row IfThen Next
        InstrOnExcel instrOnExcel = TestInstrBuilder.CreateInstrOnExcelFileName("file", instrIfThenElse);
        program.ListInstr.Add(instrOnExcel);

        //--create the compiled script -> the program
        //Program programScript = new ProgramScript(script, listInstr);

        //--create the program runner
        ActivityLogger logger = new ActivityLogger();
        ExcelProcessorNpoi excelProcessor = new ExcelProcessorNpoi();
        ProgramExecutor programExec = new ProgramExecutor(logger, excelProcessor);
        ExecResult execResult = new ExecResult();
        bool res = programExec.Exec(execResult, program);
        Assert.IsTrue(res);

        //--check the content of excel file
        var fileStream = TestExcelChecker.OpenExcel(PathExcelFilesExec + "dataOnExcel2.xlsx");
        Assert.IsNotNull(fileStream);
        var wb = TestExcelChecker.GetWorkbook(fileStream);

        // r1, c0: 9  -> not modified
        res = TestExcelChecker.CheckCellValue(wb, 0, 1, 0, 9);
        Assert.IsTrue(res);

        // r2, c0: 10 -> modified!
        res = TestExcelChecker.CheckCellValue(wb, 0, 2, 0, 10);
        Assert.IsTrue(res);
    }

    /// <summary>
    ///
    /// file= "dataOnExcelJokerB.xlsx"
    /// OnExcel file
    ///   ForEach Row
    ///     If A.Cell >10 Then A.Cell= 10
    ///   Next
    /// End OnExcel
    /// </summary>
    [TestMethod]
    public void ExecOpenExcelFileNameVarStringJokerOk()
    {
        Program program = TestInstrBuilder.CreateProgram();

        //-->SetVar #1:   file= "data1.xslx"
        //    InstrLeft:  ObjectName: file
        //    InstrRight: OpenExcel, p="data.xlsx"

        // instr left
        InstrObjectName instrObjectName = TestInstrBuilder.CreateInstrObjectName("file");

        //instr right
        InstrValue instrValue = TestInstrBuilder.CreateValueString(AddDblQuote(PathExcelFilesExec + "dataOnExcelJokerB*.xlsx"));

        InstrSetVar instrSetVar = TestInstrBuilder.CreateInstrSetVar(instrObjectName, instrValue);
        program.ListInstr.Add(instrSetVar);

        //-->OnExcel #2:
        //--If A.Cell >10 Then A.Cell= 10
        InstrIfThenElse instrIfThenElse = TestInstrBuilder.CreateInstrIfThen("A", 1, ">", 10, "A", 1, 10);

        //--OnExcel file ForEach Row IfThen Next
        InstrOnExcel instrOnExcel = TestInstrBuilder.CreateInstrOnExcelFileName("file", instrIfThenElse);
        program.ListInstr.Add(instrOnExcel);

        //--create the program runner
        ActivityLogger logger = new ActivityLogger();
        ExcelProcessorNpoi excelProcessor = new ExcelProcessorNpoi();
        ProgramExecutor programExec = new ProgramExecutor(logger, excelProcessor);
        ExecResult execResult = new ExecResult();
        bool res = programExec.Exec(execResult, program);
        Assert.IsTrue(res);

        //--check the content of excel file
        var fileStream = TestExcelChecker.OpenExcel(PathExcelFilesExec + "dataOnExcelJokerB.xlsx");
        Assert.IsNotNull(fileStream);
        var wb = TestExcelChecker.GetWorkbook(fileStream);

        // r1, c0: 9  -> not modified
        res = TestExcelChecker.CheckCellValue(wb, 0, 1, 0, 9);
        Assert.IsTrue(res);

        // r2, c0: 10 -> modified!
        res = TestExcelChecker.CheckCellValue(wb, 0, 2, 0, 10);
        Assert.IsTrue(res);

        //--check the content of excel file  B2
        fileStream = TestExcelChecker.OpenExcel(PathExcelFilesExec + "dataOnExcelJokerB2.xlsx");
        Assert.IsNotNull(fileStream);
        wb = TestExcelChecker.GetWorkbook(fileStream);

        // r1, c0: 9  -> not modified
        res = TestExcelChecker.CheckCellValue(wb, 0, 1, 0, 9);
        Assert.IsTrue(res);

        // r2, c0: 10 -> modified!
        res = TestExcelChecker.CheckCellValue(wb, 0, 2, 0, 10);
        Assert.IsTrue(res);
    }

    /// <summary>
    /// file=SelectFiles("dataOnExcel3.xlsx")
    /// OnExcel file
    ///   ForEach Row
    ///     If A.Cell >10 Then A.Cell= 10
    ///   Next
    /// End OnExcel
    /// </summary>
    [TestMethod]
    public void ExecOpenExcelFileNameVarSelectFilesOk()
    {
        Program program = TestInstrBuilder.CreateProgram();

        //-->SetVar #1:   file= SelectFiles("dataOnExcel3.xslx")
        //    InstrRight: ObjectName: file
        //    InstrLeft:  OpenExcel, p="dataOnExcel3.xlsx"

        // instr left
        InstrObjectName instrObjectName = TestInstrBuilder.CreateInstrObjectName("file");

        // instr right
        InstrSelectFiles instrSelectFiles = TestInstrBuilder.CreateInstrSelectExcelParamString(AddDblQuote(PathExcelFilesExec + "dataOnExcel3.xlsx"));

        InstrSetVar instrSetVar = TestInstrBuilder.CreateInstrSetVar(instrObjectName, instrSelectFiles);
        program.ListInstr.Add(instrSetVar);

        //-->OnExcel #2:
        //--If A.Cell >10 Then A.Cell= 10
        InstrIfThenElse instrIfThenElse = TestInstrBuilder.CreateInstrIfThen("A", 1, ">", 10, "A", 1, 10);

        //--OnExcel file ForEach Row IfThen Next
        InstrOnExcel instrOnExcel = TestInstrBuilder.CreateInstrOnExcelFileName("file", instrIfThenElse);
        program.ListInstr.Add(instrOnExcel);

        //--create the program runner
        ActivityLogger logger = new ActivityLogger();
        ExcelProcessorNpoi excelProcessor = new ExcelProcessorNpoi();
        ProgramExecutor programExec = new ProgramExecutor(logger, excelProcessor);
        ExecResult execResult = new ExecResult();
        bool res = programExec.Exec(execResult, program);
        Assert.IsTrue(res);

        //--check the content of excel file
        var fileStream = TestExcelChecker.OpenExcel(PathExcelFilesExec + "dataOnExcel3.xlsx");
        Assert.IsNotNull(fileStream);
        var wb = TestExcelChecker.GetWorkbook(fileStream);

        // r1, c0: 9  -> not modified
        res = TestExcelChecker.CheckCellValue(wb, 0, 1, 0, 9);
        Assert.IsTrue(res);

        // r2, c0: 10 -> modified!
        res = TestExcelChecker.CheckCellValue(wb, 0, 2, 0, 10);
        Assert.IsTrue(res);
    }
}