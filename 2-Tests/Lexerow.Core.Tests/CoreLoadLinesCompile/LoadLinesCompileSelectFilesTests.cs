using Lexerow.Core.System;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Tests.CoreLoadLinesCompile;

/// <summary>
/// Test the script load lines and compile from the core.
/// Focus on SelectFiles instruction.
/// No need to have input excel file. Check the script content/structure itself.
/// </summary>
[TestClass]
public class LoadLinesCompileSelectFilesTests
{
    [TestMethod]
    public void BasicOk()
    {
        ExecResult execResult;
        LexerowCore core = new LexerowCore();

        // create a basic script
        List<string> lines = [
            "file=SelectFiles(\"mydata.xlsx\")"
            ];

        // load the script and compile it
        execResult = core.LoadLinesScript("script", lines);
        Assert.IsTrue(execResult.Result);
    }

    [TestMethod]
    public void WithOneLineCommentOk()
    {
        ExecResult execResult;
        LexerowCore core = new LexerowCore();

        // create a basic script
        List<string> lines = [
            "# blabla", 
            "file=SelectFiles(\"mydata.xlsx\")"
            ];

        // load the script and compile it
        execResult = core.LoadLinesScript("script", lines);
        Assert.IsTrue(execResult.Result);
    }

    [TestMethod]
    public void FunctionNotExistsError()
    {
        ExecResult execResult;
        LexerowCore core = new LexerowCore();

        // create a basic script
        List<string> lines = [
            "# blabla",
            "file=SelectXXX(\"mydata.xlsx\")"
            ];

        // load the script and compile it
        execResult = core.LoadLinesScript("script", lines);
        Assert.IsFalse(execResult.Result);
        Assert.AreEqual(ErrorCode.ParserTokenNotExpected, execResult.ListError[0].ErrorCode);
        Assert.AreEqual("SelectXXX", execResult.ListError[0].Param);
    }

}
