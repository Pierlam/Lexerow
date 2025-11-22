using Lexerow.Core.ScriptLoad;
using Lexerow.Core.System;
using Lexerow.Core.System.ScriptDef;

namespace Lexerow.Core.Tests.ScriptLoad;

/// <summary>
/// Test only the loader of scripts.
/// </summary>
[TestClass]
public class ScriptLoaderBasicTests
{
    /// <summary>
    /// load a script.
    /// Containes 2 lines.
    /// # basic script
    /// file= SelectFiles("data.xslx")
    /// </summary>
    [TestMethod]
    public void LoadScript1Ok()
    {
        ScriptLoader loader = new ScriptLoader();

        Result result = new Result();
        string fileName = @"15-Scripts\script1.lxrw";
        loader.LoadScriptFromFile(result, "MyScript", fileName, out Script script);

        Assert.IsTrue(result.Res);
        Assert.AreEqual(2, script.ScriptLines.Count);

        // check line 1
        Assert.AreEqual(1, script.ScriptLines[0].NumLine);
        Assert.AreEqual("# basic script", script.ScriptLines[0].Line);

        // check line 2
        Assert.AreEqual(2, script.ScriptLines[1].NumLine);
        Assert.AreEqual("file= SelectFiles(\"data.xslx\")", script.ScriptLines[1].Line);
    }

    /// <summary>
    /// </summary>
    [TestMethod]
    public void LoadEmptyScript()
    {
        ScriptLoader loader = new ScriptLoader();

        Result result = new Result();
        string fileName = @"15-Scripts\scriptEmpty.lxrw";
        loader.LoadScriptFromFile(result, "MyScript", fileName, out Script script);

        Assert.IsFalse(result.Res);
        Assert.AreEqual(1, result.ListError.Count);
        Assert.AreEqual(ErrorCode.LoadScriptFileEmpty, result.ListError[0].ErrorCode);
        Assert.AreEqual(0, script.ScriptLines.Count);
    }

    /// <summary>
    /// </summary>
    [TestMethod]
    public void ScriptFileNotFoundErr()
    {
        ScriptLoader loader = new ScriptLoader();

        Result result = new Result();
        string fileName = @"15-Scripts\notexist.lxrw";
        loader.LoadScriptFromFile(result, "MyScript", fileName, out Script script);

        Assert.IsFalse(result.Res);
        Assert.AreEqual(1, result.ListError.Count);
        Assert.AreEqual(ErrorCode.FileNotFound, result.ListError[0].ErrorCode);
    }
}