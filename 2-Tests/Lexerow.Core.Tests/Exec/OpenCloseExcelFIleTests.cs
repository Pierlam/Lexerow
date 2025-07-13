using Lexerow.Core.System;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Tests.Exec;

[TestClass]
public class OpenCloseExcelFileTests
{
    /// <summary>
    /// ok the input excel file exists.
    /// </summary>
    [TestMethod]
    public void OpenExcelOk()
    {
        LexerowCore core = new LexerowCore();

        string fileName = @"10-Files\Test2OpenClose.xlsx";

        ExecResult execResult = core.ProgBuilder.CreateInstrOpenExcel("file", fileName);
        Assert.IsTrue(execResult.Result);

        execResult = core.CompileProgram();
        Assert.IsTrue(execResult.Result);

        //--Execute all saved instruction
        execResult = core.ExecuteProgram();
        Assert.IsTrue(execResult.Result);
    }

    // OpenExcel -> n'existe pas
    [TestMethod]
    public void OpenExcelNotFoundError()
    {
        LexerowCore core = new LexerowCore();

        string fileName = @"10-Files\blablabla.xlsx";

        ExecResult execResult = core.ProgBuilder.CreateInstrOpenExcel("file", fileName);
        Assert.IsTrue(execResult.Result);

        execResult = core.CompileProgram();
        Assert.IsTrue(execResult.Result);

        //--Execute all saved instruction
        execResult = core.ExecuteProgram();
        Assert.IsFalse(execResult.Result);
        Assert.AreEqual(ErrorCode.FileNotFound, execResult.ListError[0].ErrorCode);
        Assert.AreEqual(fileName, execResult.ListError[0].Param);
    }
}
