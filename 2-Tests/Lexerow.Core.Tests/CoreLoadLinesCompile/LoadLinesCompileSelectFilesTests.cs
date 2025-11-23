using Lexerow.Core.System;

namespace Lexerow.Core.Tests.CoreLoadLinesCompile;

/// <summary>
/// Test the script load from lines and compile from the core.
/// Focus on SelectFiles instruction.
/// No need to have input excel file. Check the script content/structure itself.
/// </summary>
[TestClass]
public class LoadLinesCompileSelectFilesTests
{
    [TestMethod]
    public void BasicOk()
    {
        Result result;
        LexerowCore core = new LexerowCore();

        // create a basic script
        List<string> lines = [
            "file=SelectFiles(\"mydata.xlsx\")"
            ];

        // load the script and compile it
        result = core.LoadLinesScript("script", lines);
        Assert.IsTrue(result.Res);
    }

    [TestMethod]
    public void WithOneLineCommentOk()
    {
        Result result;
        LexerowCore core = new LexerowCore();

        // create a basic script
        List<string> lines = [
            "# blabla",
            "file=SelectFiles(\"mydata.xlsx\")"
            ];

        // load the script and compile it
        result = core.LoadLinesScript("script", lines);
        Assert.IsTrue(result.Res);
    }

    [TestMethod]
    public void FunctionNotExistsError()
    {
        Result result;
        LexerowCore core = new LexerowCore();

        // create a basic script
        List<string> lines = [
            "# blabla",
            "file=SelectXXX(\"mydata.xlsx\")"
            ];

        // load the script and compile it
        result = core.LoadLinesScript("script", lines);
        Assert.IsFalse(result.Res);
        Assert.AreEqual(ErrorCode.ParserTokenNotExpected, result.ListError[0].ErrorCode);
        Assert.AreEqual("SelectXXX", result.ListError[0].Param);
    }
}