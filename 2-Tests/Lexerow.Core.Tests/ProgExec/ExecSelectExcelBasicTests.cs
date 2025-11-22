using Lexerow.Core.ExcelLayer;
using Lexerow.Core.ProgExec;
using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;
using Lexerow.Core.Tests._05_Common;
using Lexerow.Core.Tests.Common;

namespace Lexerow.Core.Tests.ProgExec;

/// <summary>
/// Test execution of the instruction SelectFiles().
/// </summary>
[TestClass]
public class ExecSelectExcelBasicTests : BaseTests
{
    /// <summary>
    /// file= SelectFiles("data1.xslx")
    /// Goal: scan files, build the final list of files.
    /// </summary>
    [TestMethod]
    public void ExecSelectFilesFilenameOk()
    {
        Program program = TestInstrBuilder.CreateProgram();

        //--SetVar #1:   file= OpenExcel("data1.xslx")
        //    InstrLeft:  ObjectName: file
        //    InstrRight: OpenExcel, p="data.xlsx"

        //-instr left
        InstrObjectName instrObjectName = TestInstrBuilder.CreateInstrObjectName("file");

        //-instr right
        InstrSelectFiles instrSelectFiles = TestInstrBuilder.CreateInstrSelectExcelParamString(AddDblQuote(PathExcelFilesExec + "data1.xlsx"));

        //-Setvar
        InstrSetVar instrSetVar = TestInstrBuilder.CreateInstrSetVar(instrObjectName, instrSelectFiles);
        program.ListInstr.Add(instrSetVar);

        //--create the program runner
        ProgramExecutor programExec = new ProgramExecutor(new ActivityLogger(), new ExcelProcessorNpoi());
        Result result = new Result();
        bool res = programExec.Exec(result, program);
        Assert.IsTrue(res);

        // found one file
        Assert.AreEqual(1, instrSelectFiles.ListSelectedFilename.Count);
        Assert.AreEqual((instrSelectFiles.ListInstrParams[0] as InstrValue).RawValue, (instrSelectFiles.ListSelectedFilename[0].InstrBase as InstrValue).RawValue);
        Assert.IsNotNull(1, instrSelectFiles.ListSelectedFilename[0].Filename);
    }

    /// <summary>
    /// file= SelectFiles("blabla.xlsx")  ->no file matching the name
    /// </summary>
    [TestMethod]
    public void ExecSelectFilesFilenameOkButNotFound()
    {
        Program program = TestInstrBuilder.CreateProgram();

        //--SetVar #1:
        //      InstrLeft:   ObjectName: file
        //      InstrRight:  OpenExcel, p="data.xlsx"

        //-instr left
        InstrObjectName instrObjectName = TestInstrBuilder.CreateInstrObjectName("file");

        //-instr right
        InstrSelectFiles instrSelectFiles = TestInstrBuilder.CreateInstrSelectExcelParamString(AddDblQuote("blabla.xlsx"));

        //-Setvar
        InstrSetVar instrSetVar = TestInstrBuilder.CreateInstrSetVar(instrObjectName, instrSelectFiles);
        program.ListInstr.Add(instrSetVar);

        //--create the program runner
        ProgramExecutor programExec = new ProgramExecutor(new ActivityLogger(), new ExcelProcessorNpoi());
        Result result = new Result();
        bool res = programExec.Exec(result, program);

        Assert.IsTrue(res);
        Assert.AreEqual(0, instrSelectFiles.ListSelectedFilename.Count);
    }

    /// <summary>
    /// name="dataName.xslx"
    /// file=SelectFiles(name)
    /// </summary>
    [TestMethod]
    public void ExecSelectFilesVarNameOk()
    {
        Program program = TestInstrBuilder.CreateProgram();

        //--SetVar #1:
        //      InstrLeft:   ObjectName: name
        //      InstrRight:  ConstValue: "data.xlsx"

        //-instr left
        InstrObjectName instrObjectName = TestInstrBuilder.CreateInstrObjectName("name");

        //-instr right
        InstrValue instrValue = TestInstrBuilder.CreateValueString(AddDblQuote(PathExcelFilesExec + "dataName.xlsx"));

        //-Setvar
        InstrSetVar instrSetVar = TestInstrBuilder.CreateInstrSetVar(instrObjectName, instrValue);
        program.ListInstr.Add(instrSetVar);

        //--SetVar #2:
        //      InstrLeft:  ObjectName: file
        //      InstRight:  SelectFiles, p=name

        //-instr left
        instrObjectName = TestInstrBuilder.CreateInstrObjectName("file");

        //-instr right
        InstrSelectFiles instrSelectFiles = TestInstrBuilder.CreateInstrSelectExcelParamObjectName("name");

        //-Setvar
        instrSetVar = TestInstrBuilder.CreateInstrSetVar(instrObjectName, instrSelectFiles);
        program.ListInstr.Add(instrSetVar);

        //--create the program runner
        ProgramExecutor programExec = new ProgramExecutor(new ActivityLogger(), new ExcelProcessorNpoi());
        Result result = new Result();
        bool res = programExec.Exec(result, program);
        Assert.IsTrue(res);

        // found one file: "....\dataName.xlsx"
        Assert.AreEqual(1, instrSelectFiles.ListSelectedFilename.Count);
        Assert.IsNotNull(1, instrSelectFiles.ListSelectedFilename[0].Filename);
        Assert.IsTrue(instrSelectFiles.ListSelectedFilename[0].Filename.EndsWith("dataName.xlsx"));
    }

    /// <summary>
    /// f="dataName.xslx"
    /// name=f
    /// file=SelectFiles(name)
    /// </summary>
    [TestMethod]
    public void ExecSelectFilesVarNameVarNameOk()
    {
        Program program = TestInstrBuilder.CreateProgram();

        //--SetVar #1:
        //      InstrLeft:   ObjectName: f
        //      InstrRight:  ConstValue: "data.xlsx"

        //-instr left
        InstrObjectName instrObjectName = TestInstrBuilder.CreateInstrObjectName("f");

        //-instr right
        InstrValue instrValue = TestInstrBuilder.CreateValueString(AddDblQuote(PathExcelFilesExec + "dataName.xlsx"));

        //-Setvar
        InstrSetVar instrSetVar = TestInstrBuilder.CreateInstrSetVar(instrObjectName, instrValue);
        program.ListInstr.Add(instrSetVar);

        //--SetVar #2: name=f
        //      InstrLeft:   ObjectName: f
        //      InstrRight:  ObjectName: name

        //-instr left
        instrObjectName = TestInstrBuilder.CreateInstrObjectName("name");

        //-instr right
        var instrObjectName2 = TestInstrBuilder.CreateInstrObjectName("f");

        //-Setvar
        instrSetVar = TestInstrBuilder.CreateInstrSetVar(instrObjectName, instrObjectName2);
        program.ListInstr.Add(instrSetVar);

        //--SetVar #3:
        //      InstrLeft:   ObjectName: file
        //      InstrRight:  SelectFiles, p=name

        //-instr left
        instrObjectName = TestInstrBuilder.CreateInstrObjectName("file");

        //-instr right
        InstrSelectFiles instrSelectFiles = TestInstrBuilder.CreateInstrSelectExcelParamObjectName("name");

        //-Setvar
        instrSetVar = TestInstrBuilder.CreateInstrSetVar(instrObjectName, instrSelectFiles);
        program.ListInstr.Add(instrSetVar);

        //--create the program runner
        ProgramExecutor programExec = new ProgramExecutor(new ActivityLogger(), new ExcelProcessorNpoi());
        Result result = new Result();
        bool res = programExec.Exec(result, program);
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
    public void ExecSelectFilesByVarNameError()
    {
        Program program = TestInstrBuilder.CreateProgram();

        //  -SetVar:
        //      InstrRight: ObjectName: file
        //      InstrLeft:  SelectFiles, p=name

        //-instr left
        var instrObjectName = TestInstrBuilder.CreateInstrObjectName("file");

        //-instr right
        InstrSelectFiles instrOpenExcel = TestInstrBuilder.CreateInstrSelectExcelParamObjectName("name");

        //-Setvar
        var instrSetVar = TestInstrBuilder.CreateInstrSetVar(instrObjectName, instrOpenExcel);
        program.ListInstr.Add(instrSetVar);

        //--create the program runner
        ProgramExecutor programExec = new ProgramExecutor(new ActivityLogger(), new ExcelProcessorNpoi());
        Result result = new Result();
        bool res = programExec.Exec(result, program);

        Assert.IsFalse(res);
        Assert.AreEqual(ErrorCode.ExecInstrVarNotFound, result.ListError[0].ErrorCode);
    }
}