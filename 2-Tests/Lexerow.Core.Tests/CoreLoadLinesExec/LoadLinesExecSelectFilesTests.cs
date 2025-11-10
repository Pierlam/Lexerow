using Lexerow.Core.System;
using Lexerow.Core.Tests.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
        ExecResult execResult;
        LexerowCore core = new LexerowCore();

        // create a basic script
        List<string> lines = [
            "file=SelectFiles(" + AddDblQuote(PathExcelFilesExec + "datLinesRunSelect1.xlsx") +")"
            ];

        // load the script and compile it
        execResult = core.LoadLinesScript("script", lines);
        Assert.IsTrue(execResult.Result);

        execResult = core.ExecuteScript("script");
        Assert.IsTrue(execResult.Result);
    }

    [TestMethod]
    public void LoadExecuteOk()
    {
        ExecResult execResult;
        LexerowCore core = new LexerowCore();

        // create a basic script
        List<string> lines = [
            "file=SelectFiles(" + AddDblQuote(PathExcelFilesExec + "datLinesRunSelect2.xlsx") +")"
            ];

        // load the script, compile it and execute it
        execResult = core.LoadExecLinesScript("script", lines);
        Assert.IsTrue(execResult.Result);
    }

}
