using Lexerow.Core.System;
using Lexerow.Core.Tests.Common;

namespace Lexerow.Core.Tests.CoreLoadFileCompile;

/// <summary>
/// Test the script load from lines and compile from the core.
/// Focus on OnExcel instruction.
/// No need to have input excel file. Check the script content/structure itself.
/// </summary>
[TestClass]
public class LoadFileCompileOnExcelTests : BaseTests
{
    [TestMethod]
    public void OnExcelBasicOk()
    {
        Result result;
        LexerowCore core = new LexerowCore();
        string scriptfile = PathScriptFiles + "loadOnExcel1.lxrw";

        // load the script and compile it
        result = core.LoadScript("script", scriptfile);
        Assert.IsTrue(result.Res);
    }
}