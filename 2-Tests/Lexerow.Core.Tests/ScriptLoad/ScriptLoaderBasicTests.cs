using Lexerow.Core.Scripts;
using Lexerow.Core.System;
using Lexerow.Core.System.Compilator;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Tests.ScriptLoad;

[TestClass]
public class ScriptLoaderBasicTests
{
    /// <summary>
    /// load a script.
    /// Containes 2 lines.
    /// # basic script
    /// file= OpenExcel("data.xslx")
    /// </summary>
    [TestMethod]
    public void LoadScript1Ok()
    {
        ScriptLoader loader = new ScriptLoader();

        ExecResult execResult = new ExecResult();
        string fileName = @"15-Scripts\script1.lxrw";
        loader.LoadScriptFromFile(execResult, fileName, out Script sourceScript);

        Assert.IsTrue(execResult.Result);
        Assert.AreEqual(2, sourceScript.ScriptLines.Count);

        // check line 1
        Assert.AreEqual(1, sourceScript.ScriptLines[0].NumLine);
        Assert.AreEqual("# basic script", sourceScript.ScriptLines[0].Line);

        // check line 2
        Assert.AreEqual(2, sourceScript.ScriptLines[1].NumLine);
        Assert.AreEqual("file= OpenExcel(\"data.xslx\")", sourceScript.ScriptLines[1].Line);

    }

    /// <summary>
    /// </summary>
    [TestMethod]
    public void LoadEmptyScript()
    {
        ScriptLoader loader = new ScriptLoader();

        ExecResult execResult = new ExecResult();
        string fileName = @"15-Scripts\scriptEmpty.lxrw";
        loader.LoadScriptFromFile(execResult, fileName, out Script sourceScript);

        Assert.IsFalse(execResult.Result);
        Assert.AreEqual(1, execResult.ListError.Count);
        Assert.AreEqual(ErrorCode.LoadScriptFileEmpty, execResult.ListError[0].ErrorCode);
        Assert.AreEqual(0, sourceScript.ScriptLines.Count);

    }

    /// <summary>
    /// </summary>
    [TestMethod]
    public void ScriptFileNotFoundErr()
    {
        ScriptLoader loader = new ScriptLoader();

        ExecResult execResult = new ExecResult();
        string fileName = @"15-Scripts\notexist.lxrw";
        loader.LoadScriptFromFile(execResult, fileName, out Script sourceScript);

        Assert.IsFalse(execResult.Result);
        Assert.AreEqual(1, execResult.ListError.Count);
        Assert.AreEqual(ErrorCode.FileNotFound, execResult.ListError[0].ErrorCode);
    }

}
