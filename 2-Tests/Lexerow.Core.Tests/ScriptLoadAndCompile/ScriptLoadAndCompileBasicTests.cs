using Lexerow.Core.System;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Tests.ScriptLoadAndCompile;

[TestClass]
public class ScriptLoadAndCompileBasicTests
{
    [TestMethod]
    public void LoadFromLinesBasicOk()
    {
        ExecResult execResult;
        LexerowCore core = new LexerowCore();

        // create a basic script
        List<string> lines = new List<string>();
        lines.Add("# blabla");
        lines.Add("file=OpenExcel(\"mydata.xlsx\")");

        //ici();

        // load the script and compile it
        execResult = core.LoadScriptFromLines("script", lines);
        Assert.IsTrue(execResult.Result);
    }

}
