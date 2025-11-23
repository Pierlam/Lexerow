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
        Result result;
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
        result = core.LoadLinesScript("script", lines);
        Assert.IsTrue(result.Res);
    }

    [TestMethod]
    public void IfACellGreaterEqualOk()
    {
        Result result;
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
        result = core.LoadLinesScript("script", lines);
        Assert.IsTrue(result.Res);
    }

    [TestMethod]
    public void IfACellDiffOk()
    {
        Result result;
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
        result = core.LoadLinesScript("script", lines);
        Assert.IsTrue(result.Res);
    }

    [TestMethod]
    public void IfACellEqualStringOk()
    {
        Result result;
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
        result = core.LoadLinesScript("script", lines);
        Assert.IsTrue(result.Res);
    }

    [TestMethod]
    public void IfACellEqualStringThenACellEqualStringOk()
    {
        Result result;
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
        result = core.LoadLinesScript("script", lines);
        Assert.IsTrue(result.Res);
    }

    [TestMethod]
    public void IfACellGreaterBCellStringOk()
    {
        Result result;
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
        result = core.LoadLinesScript("script", lines);
        Assert.IsTrue(result.Res);
    }

    [TestMethod]
    public void TwoIfThenOk()
    {
        Result result;
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
        result = core.LoadLinesScript("script", lines);
        Assert.IsTrue(result.Res);
    }

    /// <summary>
    /// Special case, ForEachRow is allowed like ForEach Row.
    /// </summary>
    [TestMethod]
    public void ForEachRowOk()
    {
        Result result;
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
        result = core.LoadLinesScript("script", lines);

        // TODO: acccept both syntax?? ForEach Row and ForEachRow
        Assert.IsTrue(result.Res);
    }

    [TestMethod]
    public void NextMissingError()
    {
        Result result;
        LexerowCore core = new LexerowCore();

        // create a basic script
        List<string> lines = [
            "OnExcel \"mydata.xlsx\"",
            "  ForEach Row",
            "    If A.Cell>10 Then A.Cell=10",
            "End OnExcel"
            ];

        // load the script and compile it
        result = core.LoadLinesScript("script", lines);
        Assert.IsFalse(result.Res);
    }

    [TestMethod]
    public void EndOnExcelMissingError()
    {
        Result result;
        LexerowCore core = new LexerowCore();

        // create a basic script
        List<string> lines = [
            "OnExcel \"mydata.xlsx\"",
            "  ForEach Row",
            "    If A.Cell>10 Then A.Cell=10",
            "Next"
            ];

        // load the script and compile it
        result = core.LoadLinesScript("script", lines);
        Assert.IsFalse(result.Res);
    }

    // test script with error: EndOnExcel in one word, ...
}