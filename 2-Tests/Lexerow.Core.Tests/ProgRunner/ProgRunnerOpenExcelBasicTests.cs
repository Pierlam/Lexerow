using Lexerow.Core.ExcelLayer;
using Lexerow.Core.ProgRun;
using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;
using Lexerow.Core.System.Compilator;
using Lexerow.Core.Tests._05_Common;
using Lexerow.Core.Tests.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Tests.ProgRunner;

/// <summary>
/// Test program runner.
/// </summary>
[TestClass]
public class ProgRunnerOpenExcelBasicTests: BaseTests
{
    /// <summary>
    /// file= OpenExcel("data1.xslx")  ->ok, file exists
    /// </summary>
    [TestMethod]
    public void RunOpenExcelOk()
    {
        Script script = new Script("scriptName", "fileName");
        List<InstrBase> listInstr = new List<InstrBase>();

        //--SetVar:   file= OpenExcel("data1.xslx")
        //    InstrRight: ObjectName: file
        //    InstrLeft:  OpenExcel, p="data.xlsx" 

        //-instr left
        InstrObjectName instrObjectName = TestInstrBuilder.BuildInstrObjectName("file");

        //-instr right
        InstrOpenExcel instrOpenExcel = TestInstrBuilder.BuildInstrOpenExcelParamString(AddDblQuote(PathExcelFilesRun + "data1.xlsx"));

        //-Setvar
        InstrSetVar instrSetVar = TestInstrBuilder.BuildInstrSetVar(instrObjectName, instrOpenExcel);
        listInstr.Add(instrSetVar);

        //--create the compiled script -> the program
        ProgramScript programScript = new ProgramScript(script, listInstr);

        //--create the program runner
        ActivityLogger logger= new ActivityLogger();
        ExcelProcessorNpoi excelProcessor= new ExcelProcessorNpoi();
        ProgramRunner programRunner = new ProgramRunner(logger,excelProcessor);
        ExecResult execResult = new ExecResult();
        bool res=programRunner.Run(execResult, programScript);
        Assert.IsTrue(res);
    }

    /// <summary>
    /// file= OpenExcel("blabla.xslx")  ->error, file not found
    /// </summary>
    [TestMethod]
    public void RunOpenExcelFileNotFound()
    {
        Script script = new Script("scriptName", "fileName");
        List<InstrBase> listInstr = new List<InstrBase>();

        //  -SetVar:
        //      InstrRight: ObjectName: file
        //      InstrLeft:  OpenExcel, p="data.xlsx" 

        //-instr left
        InstrObjectName instrObjectName = TestInstrBuilder.BuildInstrObjectName("file");

        //-instr right
        InstrOpenExcel instrOpenExcel = TestInstrBuilder.BuildInstrOpenExcelParamString(AddDblQuote("blabla.xlsx"));

        //-Setvar
        InstrSetVar instrSetVar = TestInstrBuilder.BuildInstrSetVar(instrObjectName, instrOpenExcel);
        listInstr.Add(instrSetVar);

        //--create the compiled script -> the program
        ProgramScript programScript = new ProgramScript(script, listInstr);

        //--create the program runner
        ActivityLogger logger = new ActivityLogger();
        ExcelProcessorNpoi excelProcessor = new ExcelProcessorNpoi();
        ProgramRunner programRunner = new ProgramRunner(logger, excelProcessor);
        ExecResult execResult = new ExecResult();
        bool res = programRunner.Run(execResult, programScript);

        Assert.IsFalse(res);
        Assert.AreEqual(ErrorCode.FileNotFound, execResult.ListError[0].ErrorCode);
    }

    /// <summary>
    /// name="dataName.xslx"
    /// file=OpenExcel(name)
    /// </summary>
    [TestMethod]
    public void RunOpenExcelFileByNameOk()
    {
        Script script = new Script("scriptName", "fileName");
        List<InstrBase> listInstr = new List<InstrBase>();

        //  -SetVar:
        //      InstrRight: ObjectName: name
        //      InstrLeft:  ConstValue: "data.xlsx" 

        //-instr left
        InstrObjectName instrObjectName = TestInstrBuilder.BuildInstrObjectName("name");

        //-instr right
        InstrConstValue instrConstValue = TestInstrBuilder.BuildInstrConstValueString(AddDblQuote(PathExcelFilesRun + "dataName.xlsx"));

        //-Setvar
        InstrSetVar instrSetVar = TestInstrBuilder.BuildInstrSetVar(instrObjectName, instrConstValue);
        listInstr.Add(instrSetVar);


        //  -SetVar:
        //      InstrRight: ObjectName: file
        //      InstrLeft:  OpenExcel, p=name 

        //-instr left
        instrObjectName = TestInstrBuilder.BuildInstrObjectName("file");

        //-instr right
        InstrOpenExcel instrOpenExcel = TestInstrBuilder.BuildInstrOpenExcelParamObjectName("name");

        //-Setvar
        instrSetVar = TestInstrBuilder.BuildInstrSetVar(instrObjectName, instrOpenExcel);
        listInstr.Add(instrSetVar);

        //--create the compiled script -> the program
        ProgramScript programScript = new ProgramScript(script, listInstr);

        //--create the program runner
        ActivityLogger logger = new ActivityLogger();
        ExcelProcessorNpoi excelProcessor = new ExcelProcessorNpoi();
        ProgramRunner programRunner = new ProgramRunner(logger, excelProcessor);
        ExecResult execResult = new ExecResult();
        bool res = programRunner.Run(execResult, programScript);
        Assert.IsTrue(res);
    }

    /// <summary>
    /// file=OpenExcel(name)
    /// Error: var not found!
    /// </summary>
    [TestMethod]
    public void RunOpenExcelFileByNameError()
    {
        Script script = new Script("scriptName", "fileName");
        List<InstrBase> listInstr = new List<InstrBase>();

        //  -SetVar:
        //      InstrRight: ObjectName: file
        //      InstrLeft:  OpenExcel, p=name 

        //-instr left
        var instrObjectName = TestInstrBuilder.BuildInstrObjectName("file");

        //-instr right
        InstrOpenExcel instrOpenExcel = TestInstrBuilder.BuildInstrOpenExcelParamObjectName("name");

        //-Setvar
        var instrSetVar = TestInstrBuilder.BuildInstrSetVar(instrObjectName, instrOpenExcel);
        listInstr.Add(instrSetVar);

        //--create the compiled script -> the program
        ProgramScript programScript = new ProgramScript(script, listInstr);

        //--create the program runner
        ActivityLogger logger = new ActivityLogger();
        ExcelProcessorNpoi excelProcessor = new ExcelProcessorNpoi();
        ProgramRunner programRunner = new ProgramRunner(logger, excelProcessor);
        ExecResult execResult = new ExecResult();
        bool res = programRunner.Run(execResult, programScript);

        Assert.IsFalse(res);
        Assert.AreEqual(ErrorCode.RunInstrVarNotFound, execResult.ListError[0].ErrorCode);
    }
}
