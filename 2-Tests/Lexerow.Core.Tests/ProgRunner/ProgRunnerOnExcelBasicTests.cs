using FakeItEasy;
using Lexerow.Core.ExcelLayer;
using Lexerow.Core.ProgExec;
using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;
using Lexerow.Core.System.ScriptDef;
using Lexerow.Core.Tests._05_Common;
using Lexerow.Core.Tests._20_Utils;
using Lexerow.Core.Tests.Common;
using NPOI.SS.Formula.Functions;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Net.Mime.MediaTypeNames;

namespace Lexerow.Core.Tests.ProgRunner;

[TestClass]
public class ProgRunnerOnExcelBasicTests : BaseTests
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
    public void RunOnExcelFilenameStringOk()
    {
        Script script = new Script("scriptName", "fileName");
        List<InstrBase> listInstr = new List<InstrBase>();

        //--OnExcel #1: 
        //--If A.Cell >10 Then A.Cell= 10
        InstrIfThenElse instrIfThenElse = TestInstrBuilder.CreateInstrIfThen("A", 1, ">", 10, "A", 1, 10);

        //--OnExcel "dataOnExcel1.xlsx" ForEach Row IfThen Next
        InstrOnExcel instrOnExcel = TestInstrBuilder.CreateInstrOnExcelFileString(AddDblQuote(PathExcelFilesExec + "dataOnExcel1.xlsx"), instrIfThenElse);
        listInstr.Add(instrOnExcel);

        //--create the compiled script -> the program
        ProgramScript programScript = new ProgramScript(script, listInstr);

        //--create the program runner
        ActivityLogger logger = new ActivityLogger();
        ExcelProcessorNpoi excelProcessor = new ExcelProcessorNpoi();
        ProgramExecutor programRunner = new ProgramExecutor(logger, excelProcessor);
        ExecResult execResult = new ExecResult();
        bool res = programRunner.Exec(execResult, programScript);
        Assert.IsTrue(res);

        //--check the content of excel file
        var fileStream= ExcelTestChecker.OpenExcel(PathExcelFilesExec + "dataOnExcel1.xlsx");
        Assert.IsNotNull(fileStream);
        var wb= ExcelTestChecker.GetWorkbook(fileStream);

        // r1, c0: 9  -> not modified
        res = ExcelTestChecker.CheckCellValue(wb, 0, 1, 0, 9);
        Assert.IsTrue(res);

        // r2, c0: 10 -> modified!
        res = ExcelTestChecker.CheckCellValue(wb, 0,2,0,10);
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
    public void RunOnExcelFilenameJokerStringOk()
    {
        Script script = new Script("scriptName", "fileName");
        List<InstrBase> listInstr = new List<InstrBase>();

        //--OnExcel #1: 
        //--If A.Cell >10 Then A.Cell= 10
        InstrIfThenElse instrIfThenElse = TestInstrBuilder.CreateInstrIfThen("A", 1, ">", 10, "A", 1, 10);

        //--OnExcel "dataOnExcel1.xlsx" ForEach Row IfThen Next
        InstrOnExcel instrOnExcel = TestInstrBuilder.CreateInstrOnExcelFileString(AddDblQuote(PathExcelFilesExec + "dataOnExcelJokerA*.xlsx"), instrIfThenElse);
        listInstr.Add(instrOnExcel);

        //--create the compiled script -> the program
        ProgramScript programScript = new ProgramScript(script, listInstr);

        //--create the program runner
        ActivityLogger logger = new ActivityLogger();
        ExcelProcessorNpoi excelProcessor = new ExcelProcessorNpoi();
        ProgramExecutor programRunner = new ProgramExecutor(logger, excelProcessor);
        ExecResult execResult = new ExecResult();
        bool res = programRunner.Exec(execResult, programScript);
        Assert.IsTrue(res);

        //--check the content of excel file
        var fileStream = ExcelTestChecker.OpenExcel(PathExcelFilesExec + "dataOnExcelJokerA.xlsx");
        Assert.IsNotNull(fileStream);
        var wb = ExcelTestChecker.GetWorkbook(fileStream);

        // r1, c0: 9  -> not modified
        res = ExcelTestChecker.CheckCellValue(wb, 0, 1, 0, 9);
        Assert.IsTrue(res);

        // r2, c0: 10 -> modified!
        res = ExcelTestChecker.CheckCellValue(wb, 0, 2, 0, 10);
        Assert.IsTrue(res);

        //--check the content of excel file
        fileStream = ExcelTestChecker.OpenExcel(PathExcelFilesExec + "dataOnExcelJokerA2.xlsx");
        Assert.IsNotNull(fileStream);
        wb = ExcelTestChecker.GetWorkbook(fileStream);

        // r1, c0: 9  -> not modified
        res = ExcelTestChecker.CheckCellValue(wb, 0, 1, 0, 9);
        Assert.IsTrue(res);

        // r2, c0: 10 -> modified!
        res = ExcelTestChecker.CheckCellValue(wb, 0, 2, 0, 10);
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
    public void RunOpenExcelFileNameVarStringOk()
    {
        Script script = new Script("scriptName", "fileName");
        List<InstrBase> listInstr = new List<InstrBase>();

        //-->SetVar #1:   file= "data1.xslx"
        //    InstrLeft:  ObjectName: file
        //    InstrRight: OpenExcel, p="data.xlsx" 

        // instr left
        InstrObjectName instrObjectName = TestInstrBuilder.BuildInstrObjectName("file");

        //instr right
        InstrConstValue instrConstValue = TestInstrBuilder.BuildInstrConstValueString(AddDblQuote(PathExcelFilesExec + "dataOnExcel2.xlsx"));

        InstrSetVar instrSetVar = TestInstrBuilder.BuildInstrSetVar(instrObjectName, instrConstValue);
        listInstr.Add(instrSetVar);

        //-->OnExcel #2: 
        //--If A.Cell >10 Then A.Cell= 10
        InstrIfThenElse instrIfThenElse = TestInstrBuilder.CreateInstrIfThen("A", 1, ">", 10, "A", 1, 10);

        //--OnExcel file ForEach Row IfThen Next
        InstrOnExcel instrOnExcel = TestInstrBuilder.CreateInstrOnExcelFileName("file", instrIfThenElse);
        listInstr.Add(instrOnExcel);

        //--create the compiled script -> the program
        ProgramScript programScript = new ProgramScript(script, listInstr);

        //--create the program runner
        ActivityLogger logger = new ActivityLogger();
        ExcelProcessorNpoi excelProcessor = new ExcelProcessorNpoi();
        ProgramExecutor programRunner = new ProgramExecutor(logger, excelProcessor);
        ExecResult execResult = new ExecResult();
        bool res = programRunner.Exec(execResult, programScript);
        Assert.IsTrue(res);

        //--check the content of excel file
        var fileStream = ExcelTestChecker.OpenExcel(PathExcelFilesExec + "dataOnExcel2.xlsx");
        Assert.IsNotNull(fileStream);
        var wb = ExcelTestChecker.GetWorkbook(fileStream);

        // r1, c0: 9  -> not modified
        res = ExcelTestChecker.CheckCellValue(wb, 0, 1, 0, 9);
        Assert.IsTrue(res);

        // r2, c0: 10 -> modified!
        res = ExcelTestChecker.CheckCellValue(wb, 0, 2, 0, 10);
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
    public void RunOpenExcelFileNameVarStringJokerOk()
    {
        Script script = new Script("scriptName", "fileName");
        List<InstrBase> listInstr = new List<InstrBase>();

        //-->SetVar #1:   file= "data1.xslx"
        //    InstrLeft:  ObjectName: file
        //    InstrRight: OpenExcel, p="data.xlsx" 

        // instr left
        InstrObjectName instrObjectName = TestInstrBuilder.BuildInstrObjectName("file");

        //instr right
        InstrConstValue instrConstValue = TestInstrBuilder.BuildInstrConstValueString(AddDblQuote(PathExcelFilesExec + "dataOnExcelJokerB*.xlsx"));

        InstrSetVar instrSetVar = TestInstrBuilder.BuildInstrSetVar(instrObjectName, instrConstValue);
        listInstr.Add(instrSetVar);

        //-->OnExcel #2: 
        //--If A.Cell >10 Then A.Cell= 10
        InstrIfThenElse instrIfThenElse = TestInstrBuilder.CreateInstrIfThen("A", 1, ">", 10, "A", 1, 10);

        //--OnExcel file ForEach Row IfThen Next
        InstrOnExcel instrOnExcel = TestInstrBuilder.CreateInstrOnExcelFileName("file", instrIfThenElse);
        listInstr.Add(instrOnExcel);

        //--create the compiled script -> the program
        ProgramScript programScript = new ProgramScript(script, listInstr);

        //--create the program runner
        ActivityLogger logger = new ActivityLogger();
        ExcelProcessorNpoi excelProcessor = new ExcelProcessorNpoi();
        ProgramExecutor programRunner = new ProgramExecutor(logger, excelProcessor);
        ExecResult execResult = new ExecResult();
        bool res = programRunner.Exec(execResult, programScript);
        Assert.IsTrue(res);

        //--check the content of excel file
        var fileStream = ExcelTestChecker.OpenExcel(PathExcelFilesExec + "dataOnExcelJokerB.xlsx");
        Assert.IsNotNull(fileStream);
        var wb = ExcelTestChecker.GetWorkbook(fileStream);

        // r1, c0: 9  -> not modified
        res = ExcelTestChecker.CheckCellValue(wb, 0, 1, 0, 9);
        Assert.IsTrue(res);

        // r2, c0: 10 -> modified!
        res = ExcelTestChecker.CheckCellValue(wb, 0, 2, 0, 10);
        Assert.IsTrue(res);


        //--check the content of excel file  B2
        fileStream = ExcelTestChecker.OpenExcel(PathExcelFilesExec + "dataOnExcelJokerB2.xlsx");
        Assert.IsNotNull(fileStream);
        wb = ExcelTestChecker.GetWorkbook(fileStream);

        // r1, c0: 9  -> not modified
        res = ExcelTestChecker.CheckCellValue(wb, 0, 1, 0, 9);
        Assert.IsTrue(res);

        // r2, c0: 10 -> modified!
        res = ExcelTestChecker.CheckCellValue(wb, 0, 2, 0, 10);
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
    public void RunOpenExcelFileNameVarSelectFilesOk()
    {
        Script script = new Script("scriptName", "fileName");
        List<InstrBase> listInstr = new List<InstrBase>();

        //-->SetVar #1:   file= SelectFiles("dataOnExcel3.xslx")
        //    InstrRight: ObjectName: file
        //    InstrLeft:  OpenExcel, p="dataOnExcel3.xlsx" 

        // instr left
        InstrObjectName instrObjectName = TestInstrBuilder.BuildInstrObjectName("file");

        // instr right
        InstrSelectFiles instrSelectFiles = TestInstrBuilder.BuildInstrSelectExcelParamString(AddDblQuote(PathExcelFilesExec + "dataOnExcel3.xlsx"));

        InstrSetVar instrSetVar = TestInstrBuilder.BuildInstrSetVar(instrObjectName, instrSelectFiles);
        listInstr.Add(instrSetVar);

        //-->OnExcel #2: 
        //--If A.Cell >10 Then A.Cell= 10
        InstrIfThenElse instrIfThenElse = TestInstrBuilder.CreateInstrIfThen("A", 1, ">", 10, "A", 1, 10);

        //--OnExcel file ForEach Row IfThen Next
        InstrOnExcel instrOnExcel = TestInstrBuilder.CreateInstrOnExcelFileName("file", instrIfThenElse);
        listInstr.Add(instrOnExcel);

        //--create the compiled script -> the program
        ProgramScript programScript = new ProgramScript(script, listInstr);

        //--create the program runner
        ActivityLogger logger = new ActivityLogger();
        ExcelProcessorNpoi excelProcessor = new ExcelProcessorNpoi();
        ProgramExecutor programRunner = new ProgramExecutor(logger, excelProcessor);
        ExecResult execResult = new ExecResult();
        bool res = programRunner.Exec(execResult, programScript);
        Assert.IsTrue(res);

        //--check the content of excel file
        var fileStream = ExcelTestChecker.OpenExcel(PathExcelFilesExec + "dataOnExcel3.xlsx");
        Assert.IsNotNull(fileStream);
        var wb = ExcelTestChecker.GetWorkbook(fileStream);

        // r1, c0: 9  -> not modified
        res = ExcelTestChecker.CheckCellValue(wb, 0, 1, 0, 9);
        Assert.IsTrue(res);

        // r2, c0: 10 -> modified!
        res = ExcelTestChecker.CheckCellValue(wb, 0, 2, 0, 10);
        Assert.IsTrue(res);
    }

}
