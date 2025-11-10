using Lexerow.Core.System;
using Lexerow.Core.Tests._20_Utils;
using Lexerow.Core.Tests.Common;
using Microsoft.VisualStudio.TestPlatform.ObjectModel.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
        ExecResult execResult;
        LexerowCore core = new LexerowCore();
        string scriptfile = PathScriptFiles + "loadOnExcel1.lxrw";

        // load the script and compile it
        execResult = core.LoadScript("script", scriptfile);
        Assert.IsTrue(execResult.Result);
    }
}
