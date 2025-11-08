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
/// Test execution of the instruction SelectFiles().
/// Test program runner.
/// </summary>
[TestClass]
public class ProgRunnerSelectExcelBasicTests: BaseTests
{
    /// <summary>
    /// file= SelectFiles("data1.xslx")  
    /// Goal: scan files, build the final list of files.
    /// </summary>
    [TestMethod]
    public void RunSelectFilesFilenameOk()
    {
        Script script = new Script("scriptName", "fileName");
        List<InstrBase> listInstr = new List<InstrBase>();

        //--SetVar #1:   file= OpenExcel("data1.xslx")
        //    InstrLeft:  ObjectName: file
        //    InstrRight: OpenExcel, p="data.xlsx" 

        //-instr left
        InstrObjectName instrObjectName = TestInstrBuilder.BuildInstrObjectName("file");

        //-instr right
        InstrSelectFiles instrSelectFiles = TestInstrBuilder.BuildInstrSelectExcelParamString(AddDblQuote(PathExcelFilesRun + "data1.xlsx"));

        //-Setvar
        InstrSetVar instrSetVar = TestInstrBuilder.BuildInstrSetVar(instrObjectName, instrSelectFiles);
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

        // found one file
        Assert.AreEqual(1,instrSelectFiles.ListSelectedFilename.Count);
        Assert.AreEqual((instrSelectFiles.ListInstrParams[0] as InstrConstValue).RawValue, (instrSelectFiles.ListSelectedFilename[0].InstrBase as InstrConstValue).RawValue);
        Assert.IsNotNull(1, instrSelectFiles.ListSelectedFilename[0].Filename);
    }

    /// <summary>
    /// file= SelectFiles("blabla.xlsx")  ->no file matching the name
    /// </summary>
    [TestMethod]
    public void RunSelectFilesFilenameOkButNotFound()
    {
        Script script = new Script("scriptName", "fileName");
        List<InstrBase> listInstr = new List<InstrBase>();

        //--SetVar #1:
        //      InstrLeft:   ObjectName: file
        //      InstrRight:  OpenExcel, p="data.xlsx" 

        //-instr left
        InstrObjectName instrObjectName = TestInstrBuilder.BuildInstrObjectName("file");

        //-instr right
        InstrSelectFiles instrSelectFiles = TestInstrBuilder.BuildInstrSelectExcelParamString(AddDblQuote("blabla.xlsx"));

        //-Setvar
        InstrSetVar instrSetVar = TestInstrBuilder.BuildInstrSetVar(instrObjectName, instrSelectFiles);
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
        Assert.AreEqual(0, instrSelectFiles.ListSelectedFilename.Count);
    }

    /// <summary>
    /// name="dataName.xslx"
    /// file=SelectFiles(name)
    /// </summary>
    [TestMethod]
    public void RunSelectFilesVarNameOk()
    {
        Script script = new Script("scriptName", "fileName");
        List<InstrBase> listInstr = new List<InstrBase>();

        //--SetVar #1:
        //      InstrLeft:   ObjectName: name
        //      InstrRight:  ConstValue: "data.xlsx" 

        //-instr left
        InstrObjectName instrObjectName = TestInstrBuilder.BuildInstrObjectName("name");

        //-instr right
        InstrConstValue instrConstValue = TestInstrBuilder.BuildInstrConstValueString(AddDblQuote(PathExcelFilesRun + "dataName.xlsx"));

        //-Setvar
        InstrSetVar instrSetVar = TestInstrBuilder.BuildInstrSetVar(instrObjectName, instrConstValue);
        listInstr.Add(instrSetVar);


        //--SetVar #2:
        //      InstrLeft:  ObjectName: file
        //      InstRight:  SelectFiles, p=name 

        //-instr left
        instrObjectName = TestInstrBuilder.BuildInstrObjectName("file");

        //-instr right
        InstrSelectFiles instrSelectFiles = TestInstrBuilder.BuildInstrSelectExcelParamObjectName("name");

        //-Setvar
        instrSetVar = TestInstrBuilder.BuildInstrSetVar(instrObjectName, instrSelectFiles);
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


        // found one file: "....\dataName.xlsx"
        Assert.AreEqual(1, instrSelectFiles.ListSelectedFilename.Count);
        Assert.IsNotNull(1, instrSelectFiles.ListSelectedFilename[0].Filename);
        Assert.IsTrue(instrSelectFiles.ListSelectedFilename[0].Filename.EndsWith("dataName.xlsx"));

        //Assert.AreEqual((instrSelectFiles.ListInstrParams[0] as InstrConstValue).RawValue, (instrSelectFiles.ListFinalFilename[0].InstrBase as InstrConstValue).RawValue);
    }

    /// <summary>
    /// f="dataName.xslx"
    /// name=f
    /// file=SelectFiles(name)
    /// </summary>
    [TestMethod]
    public void RunSelectFilesVarNameVarNameOk()
    {
        Script script = new Script("scriptName", "fileName");
        List<InstrBase> listInstr = new List<InstrBase>();

        //--SetVar #1:
        //      InstrLeft:   ObjectName: f
        //      InstrRight:  ConstValue: "data.xlsx" 

        //-instr left
        InstrObjectName instrObjectName = TestInstrBuilder.BuildInstrObjectName("f");

        //-instr right
        InstrConstValue instrConstValue = TestInstrBuilder.BuildInstrConstValueString(AddDblQuote(PathExcelFilesRun + "dataName.xlsx"));

        //-Setvar
        InstrSetVar instrSetVar = TestInstrBuilder.BuildInstrSetVar(instrObjectName, instrConstValue);
        listInstr.Add(instrSetVar);

        //--SetVar #2: name=f
        //      InstrLeft:   ObjectName: f
        //      InstrRight:  ObjectName: name 

        //-instr left
        instrObjectName = TestInstrBuilder.BuildInstrObjectName("name");

        //-instr right
        var instrObjectName2 = TestInstrBuilder.BuildInstrObjectName("f");

        //-Setvar
        instrSetVar = TestInstrBuilder.BuildInstrSetVar(instrObjectName, instrObjectName2);
        listInstr.Add(instrSetVar);

        //--SetVar #3:
        //      InstrLeft:   ObjectName: file
        //      InstrRight:  SelectFiles, p=name 

        //-instr left
        instrObjectName = TestInstrBuilder.BuildInstrObjectName("file");

        //-instr right
        InstrSelectFiles instrSelectFiles = TestInstrBuilder.BuildInstrSelectExcelParamObjectName("name");

        //-Setvar
        instrSetVar = TestInstrBuilder.BuildInstrSetVar(instrObjectName, instrSelectFiles);
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


        // found one file: "....\dataName.xlsx"
        Assert.AreEqual(1, instrSelectFiles.ListSelectedFilename.Count);
        Assert.IsNotNull(1, instrSelectFiles.ListSelectedFilename[0].Filename);
        Assert.IsTrue(instrSelectFiles.ListSelectedFilename[0].Filename.EndsWith("dataName.xlsx"));
    }

    /// <summary>
    /// file=SelectFiles(name)
    /// Error: var not found!
    /// </summary>
    [TestMethod]
    public void RunSelectFilesByVarNameError()
    {
        Script script = new Script("scriptName", "fileName");
        List<InstrBase> listInstr = new List<InstrBase>();

        //  -SetVar:
        //      InstrRight: ObjectName: file
        //      InstrLeft:  SelectFiles, p=name 

        //-instr left
        var instrObjectName = TestInstrBuilder.BuildInstrObjectName("file");

        //-instr right
        InstrSelectFiles instrOpenExcel = TestInstrBuilder.BuildInstrSelectExcelParamObjectName("name");

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
