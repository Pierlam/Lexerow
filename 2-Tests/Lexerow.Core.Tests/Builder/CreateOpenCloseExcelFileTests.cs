using Lexerow.Core.System;

namespace Lexerow.Core.Tests.Builder;

/// <summary>
/// Test buidl instruction: OpenExcel and CloseExcel
/// </summary>
[TestClass]
public sealed class CreateOpenCloseExcelFileTests
{
    /// <summary>
    /// file=OpenExcel(...)
    /// </summary>
    [TestMethod]
    public void OpenExcelOk()
    {
        LexerowCore core = new LexerowCore();

        string fileName = @"..\..\..\10-Files\TestCoreOpenClose.xlsx";

        ExecResult execResult = core.ProgBuilder.CreateInstrOpenExcel("file", fileName);

        Assert.IsNotNull(execResult);
        Assert.IsTrue(execResult.Result);
    }

    /// <summary>
    /// file=OpenExcel(...)
    /// </summary>
    [TestMethod]
    public void OpenExcelFileNameNullFail()
    {
        LexerowCore core = new LexerowCore();

        //string fileName = @"..\..\..\10-Files\TestCoreOpenClose.xlsx";

        ExecResult execResult = core.ProgBuilder.CreateInstrOpenExcel("file", null);

        Assert.IsNotNull(execResult);
        Assert.IsFalse(execResult.Result);
        Assert.AreEqual(ErrorCode.FileNameNullOrEmpty, execResult.ListError[0].ErrorCode);
    }

    /// <summary>
    /// file=OpenExcel(...)
    /// </summary>
    [TestMethod]
    public void OpenExcelObjectWrong()
    {
        LexerowCore core = new LexerowCore();

        string fileName = @"..\..\..\10-Files\TestCoreOpenClose.xlsx";

        ExecResult execResult = core.ProgBuilder.CreateInstrOpenExcel("fi<!ù$=le", fileName);

        Assert.IsNotNull(execResult);
        Assert.IsFalse(execResult.Result);
        //TODO: fix it
        Assert.AreEqual(ErrorCode.ExcelFileObjectNameSyntaxWrong, execResult.ListError[0].ErrorCode);
    }


    /// <summary>
    /// file=OpenExcel(...)
    /// </summary>
    [TestMethod]
    public void OpenExcelFileNameTwoTimesFail()
    {
        LexerowCore core = new LexerowCore();

        string fileName = @"..\..\..\10-Files\TestCoreOpenClose.xlsx";

        ExecResult execResult = core.ProgBuilder.CreateInstrOpenExcel("file", fileName);

        Assert.IsNotNull(execResult);
        Assert.IsTrue(execResult.Result);

        // 2nd time!
        execResult = core.ProgBuilder.CreateInstrOpenExcel("file2", fileName);
        Assert.IsNotNull(execResult);
        Assert.IsFalse(execResult.Result);
        Assert.AreEqual(ErrorCode.ExcelFileNameAlreadyOpen, execResult.ListError[0].ErrorCode);
    }

    /// <summary>
    /// file=OpenExcel(...)
    /// </summary>
    [TestMethod]
    public void OpenExcelWithObjNameFail()
    {
        LexerowCore core = new LexerowCore();

        string fileName = @"..\..\..\10-Files\MyExcel.xlsx";

        ExecResult execResult = core.ProgBuilder.CreateInstrOpenExcel("file", fileName);

        Assert.IsNotNull(execResult);
        Assert.IsTrue(execResult.Result);

        // 2nd time!
        fileName = @"..\..\..\10-Files\MyExcel2.xlsx";
        execResult = core.ProgBuilder.CreateInstrOpenExcel("file", fileName);
        Assert.IsNotNull(execResult);
        Assert.IsFalse(execResult.Result);
        // TODO:
        Assert.AreEqual(ErrorCode.ExcelFileObjectNameAlreadyOpen, execResult.ListError[0].ErrorCode);
    }

    // close a file which does not exists

    // close a file 2 times


}
