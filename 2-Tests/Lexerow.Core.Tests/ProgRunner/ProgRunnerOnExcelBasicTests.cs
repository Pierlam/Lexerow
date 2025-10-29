using FakeItEasy;
using Lexerow.Core.ExcelLayer;
using Lexerow.Core.ProgRun;
using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;
using Lexerow.Core.System.Compilator;
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
    public void RunOpenExcelFileStringOk()
    {
        Script script = new Script("scriptName", "fileName");
        List<InstrBase> listInstr = new List<InstrBase>();

        //--If A.Cell >10 Then A.Cell= 10
        InstrIfThenElse instrIfThenElse = TestInstrBuilder.CreateInstrIfThen("A", 1, ">", 10, "A", 1, 10);

        //--ForEach Row
        InstrForEach instrForEach= TestInstrBuilder.CreateInstrForEach(instrIfThenElse);

        //--OnExcel "dataOnExcel1.xlsx"
        InstrOnExcel instrOnExcel = TestInstrBuilder.CreateInstrOnExcelFileString(AddDblQuote(PathExcelFilesRun + "dataOnExcel1.xlsx"), instrIfThenElse);
        listInstr.Add(instrOnExcel);

        //--create the compiled script -> the program
        ProgramScript programScript = new ProgramScript(script, listInstr);

        //--create the program runner
        ActivityLogger logger = new ActivityLogger();
        ExcelProcessorNpoi excelProcessor = new ExcelProcessorNpoi();
        ProgramRunner programRunner = new ProgramRunner(logger, excelProcessor);
        ExecResult execResult = new ExecResult();
        bool res = programRunner.Run(execResult, programScript);
        Assert.IsTrue(res);

        //--check the content of excel file
        var fileStream= TestExcelChecker.OpenExcel(PathExcelFilesRun + "dataOnExcel1.xlsx");
        Assert.IsNotNull(fileStream);
        var wb= TestExcelChecker.GetWorkbook(fileStream);

        // r1, c0: 9  -> not modified
        res = TestExcelChecker.CheckCellValue(wb, 0, 1, 0, 9);
        Assert.IsTrue(res);

        // r2, c0: 10 -> modified!
        res = TestExcelChecker.CheckCellValue(wb, 0,2,0,10);
        Assert.IsTrue(res);
    }

    /// <summary>
    ///     
    /// file=OpenExcel("dataOnExcel2.xlsx")
    /// OnExcel file
    ///   ForEach Row
    ///     If A.Cell >10 Then A.Cell= 10
    ///   Next
    /// End OnExcel    
    /// </summary>
    [TestMethod]
    public void RunOpenExcelFileNameOk()
    {
        Script script = new Script("scriptName", "fileName");
        List<InstrBase> listInstr = new List<InstrBase>();

        //--SetVar:   file= OpenExcel("data1.xslx")
        //    InstrRight: ObjectName: file
        //    InstrLeft:  OpenExcel, p="data.xlsx" 

        //-instr left
        InstrObjectName instrObjectName = TestInstrBuilder.BuildInstrObjectName("file");

        //-instr right
        InstrOpenExcel instrOpenExcel = TestInstrBuilder.BuildInstrOpenExcelParamString(AddDblQuote(PathExcelFilesRun + "dataOnExcel2.xlsx"));

        //-Setvar
        InstrSetVar instrSetVar = TestInstrBuilder.BuildInstrSetVar(instrObjectName, instrOpenExcel);
        listInstr.Add(instrSetVar);

        //--If A.Cell >10 Then A.Cell= 10
        InstrIfThenElse instrIfThenElse = TestInstrBuilder.CreateInstrIfThen("A", 1, ">", 10, "A", 1, 10);

        //--ForEach Row
        InstrForEach instrForEach = TestInstrBuilder.CreateInstrForEach(instrIfThenElse);

        //--OnExcel file
        InstrOnExcel instrOnExcel = TestInstrBuilder.CreateInstrOnExcelFileName("file", instrForEach);
        listInstr.Add(instrOnExcel);

        //--create the compiled script -> the program
        ProgramScript programScript = new ProgramScript(script, listInstr);

        //--create the program runner
        ActivityLogger logger = new ActivityLogger();
        ExcelProcessorNpoi excelProcessor = new ExcelProcessorNpoi();
        ProgramRunner programRunner = new ProgramRunner(logger, excelProcessor);
        ExecResult execResult = new ExecResult();
        bool res = programRunner.Run(execResult, programScript);
        Assert.IsTrue(res);

        //--check the content of excel file
        var fileStream = TestExcelChecker.OpenExcel(PathExcelFilesRun + "dataOnExcel2.xlsx");
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
