using Lexerow.Core.System;

namespace Lexerow.Core.Tests.CoreLoadLinesCompile;

/// <summary>
/// Test the script load from lines and compile from the core.
/// Focus on OnExcel instruction.
/// No need to have input excel file. Check the script content/structure itself.
/// </summary>
[TestClass]
public class LoadLinesCompileOnExcelTests
{
    [TestMethod]
    public void BasicOk()
    {
        ExecResult execResult;
        LexerowCore core = new LexerowCore();

        // create a basic script
        List<string> lines = [
            "OnExcel \"mydata.xlsx\"",
            "  ForEach Row",
            "    If A.Cell>10 Then A.Cell=10",
            "  Next",
            "End OnExcel"
            ];

        // load the script and compile it
        execResult = core.LoadLinesScript("script", lines);
        Assert.IsTrue(execResult.Result);
    }

    [TestMethod]
    public void IfACellGreaterEqualOk()
    {
        ExecResult execResult;
        LexerowCore core = new LexerowCore();

        // create a basic script
        List<string> lines = [
            "OnExcel \"mydata.xlsx\"",
            "  ForEach Row",
            "    If A.Cell>=10 Then A.Cell=10",
            "  Next",
            "End OnExcel"
            ];

        // load the script and compile it
        execResult = core.LoadLinesScript("script", lines);
        Assert.IsTrue(execResult.Result);
    }

    [TestMethod]
    public void IfACellDiffOk()
    {
        ExecResult execResult;
        LexerowCore core = new LexerowCore();

        // create a basic script
        List<string> lines = [
            "OnExcel \"mydata.xlsx\"",
            "  ForEach Row",
            "    If A.Cell<>10 Then A.Cell=10",
            "  Next",
            "End OnExcel"
            ];

        // load the script and compile it
        execResult = core.LoadLinesScript("script", lines);
        Assert.IsTrue(execResult.Result);
    }

    [TestMethod]
    public void IfACellEqualStringOk()
    {
        ExecResult execResult;
        LexerowCore core = new LexerowCore();

        // create a basic script
        List<string> lines = [
            "OnExcel \"mydata.xlsx\"",
            "  ForEach Row",
            "    If A.Cell=\"hello\" Then A.Cell=10",
            "  Next",
            "End OnExcel"
            ];

        // load the script and compile it
        execResult = core.LoadLinesScript("script", lines);
        Assert.IsTrue(execResult.Result);
    }

    [TestMethod]
    public void IfACellEqualStringThenACellEqualStringOk()
    {
        ExecResult execResult;
        LexerowCore core = new LexerowCore();

        // create a basic script
        List<string> lines = [
            "OnExcel \"mydata.xlsx\"",
            "  ForEach Row",
            "    If A.Cell=\"hello\" Then A.Cell=\"bonjour\"",
            "  Next",
            "End OnExcel"
            ];

        // load the script and compile it
        execResult = core.LoadLinesScript("script", lines);
        Assert.IsTrue(execResult.Result);
    }

    [TestMethod]
    public void IfACellGreaterBCellStringOk()
    {
        ExecResult execResult;
        LexerowCore core = new LexerowCore();

        // create a basic script
        List<string> lines = [
            "OnExcel \"mydata.xlsx\"",
            "  ForEach Row",
            "    If A.Cell>B.Cell Then C.Cell=10",
            "  Next",
            "End OnExcel"
            ];

        // load the script and compile it
        execResult = core.LoadLinesScript("script", lines);
        Assert.IsTrue(execResult.Result);
    }

    [TestMethod]
    public void TwoIfThenOk()
    {
        ExecResult execResult;
        LexerowCore core = new LexerowCore();

        // create a basic script
        List<string> lines = [
            "OnExcel \"mydata.xlsx\"",
            "  ForEach Row",
            "    If A.Cell>10 Then A.Cell=10",
            "    If B.Cell>12 Then B.Cell=12",
            "  Next",
            "End OnExcel"
            ];

        // load the script and compile it
        execResult = core.LoadLinesScript("script", lines);
        Assert.IsTrue(execResult.Result);
    }

    /// <summary>
    /// Special case, ForEachRow is allowed like ForEach Row.
    /// </summary>
    [TestMethod]
    public void ForEachRowOk()
    {
        ExecResult execResult;
        LexerowCore core = new LexerowCore();

        // create a basic script
        List<string> lines = [
            "OnExcel \"mydata.xlsx\"",
            "  ForEachRow",
            "    If A.Cell>10 Then A.Cell=10",
            "  Next",
            "End OnExcel"
            ];

        // load the script and compile it
        execResult = core.LoadLinesScript("script", lines);

        // TODO: acccept both syntax?? ForEach Row and ForEachRow
        Assert.IsTrue(execResult.Result);
    }

    [TestMethod]
    public void NextMissingError()
    {
        ExecResult execResult;
        LexerowCore core = new LexerowCore();

        // create a basic script
        List<string> lines = [
            "OnExcel \"mydata.xlsx\"",
            "  ForEach Row",
            "    If A.Cell>10 Then A.Cell=10",
            "End OnExcel"
            ];

        // load the script and compile it
        execResult = core.LoadLinesScript("script", lines);
        Assert.IsFalse(execResult.Result);
    }

    [TestMethod]
    public void EndOnExcelMissingError()
    {
        ExecResult execResult;
        LexerowCore core = new LexerowCore();

        // create a basic script
        List<string> lines = [
            "OnExcel \"mydata.xlsx\"",
            "  ForEach Row",
            "    If A.Cell>10 Then A.Cell=10",
            "Next"
            ];

        // load the script and compile it
        execResult = core.LoadLinesScript("script", lines);
        Assert.IsFalse(execResult.Result);
    }

    // test script with error: EndOnExcel in one word, ...
}