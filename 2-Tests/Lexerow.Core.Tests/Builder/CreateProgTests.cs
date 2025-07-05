using Lexerow.Core.System;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Tests.Builder;

[TestClass]
public class CreateProgTests
{
    /// <summary>
    /// file=OpenExcel(...)
    /// </summary>
    [TestMethod]
    public void CreateNewProgOk()
    {
        LexerowCore core = new LexerowCore();

        string fileName = @"..\..\..\10-Files\TestCoreOpenClose.xlsx";

        ExecResult execResult = core.ProgBuilder.CreateProgram("MyProg");
        Assert.IsNotNull(execResult);
        Assert.IsTrue(execResult.Result);

        // add instr on this new program
        execResult = core.ProgBuilder.CreateInstrOpenExcel("file", fileName);

        //Assert.IsNotNull(execResult);
        Assert.IsTrue(execResult.Result);
    }

}
