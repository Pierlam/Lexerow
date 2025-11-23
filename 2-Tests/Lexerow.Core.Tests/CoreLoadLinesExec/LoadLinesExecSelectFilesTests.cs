using Lexerow.Core.System;
using Lexerow.Core.Tests.Common;

namespace Lexerow.Core.Tests.CoreLoadLinesRun;

/// <summary>
/// Test the script load, compile and execute from the core.
/// Focus on SelectFiles instruction.
/// Need to have input excel files ready.
/// </summary>
[TestClass]
public class LoadLinesExecSelectFilesTests : BaseTests
{
    [TestMethod]
    public void LoadThenExecuteOk()
    {
        Result result;
        LexerowCore core = new LexerowCore();

        // create a basic script
        List<string> lines = [
            "file=SelectFiles(" + AddDblQuote(PathExcelFilesExec + "datLinesRunSelect1.xlsx") +")"
            ];

        // load the script and compile it
        result = core.LoadLinesScript("script", lines);
        Assert.IsTrue(result.Res);

        result = core.ExecuteScript("script");
        Assert.IsTrue(result.Res);
    }

    [TestMethod]
    public void LoadExecuteOk()
    {
        Result result;
        LexerowCore core = new LexerowCore();

        // create a basic script
        List<string> lines = [
            "file=SelectFiles(" + AddDblQuote(PathExcelFilesExec + "datLinesRunSelect2.xlsx") +")"
            ];

        // load the script, compile it and execute it
        result = core.LoadExecLinesScript("script", lines);
        Assert.IsTrue(result.Res);
    }
}