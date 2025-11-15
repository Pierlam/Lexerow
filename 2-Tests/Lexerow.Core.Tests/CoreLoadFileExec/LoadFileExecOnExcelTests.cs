using Lexerow.Core.System;
using Lexerow.Core.Tests._20_Utils;
using Lexerow.Core.Tests.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Tests.CoreLoadFileExec;

/// <summary>
/// Test the script load from txt file, compile and execute from the core.
/// Focus on OnExcel instruction.
/// Need to have script and input excel files ready.
/// </summary>
[TestClass]
public class LoadFileExecOnExcelTests : BaseTests
{
    [TestMethod]
    public void OnExcelBasicOk()
    {
        ExecResult execResult;
        LexerowCore core = new LexerowCore();
        string scriptfile = PathScriptFiles + "execOnExcel1.lxrw";

        // load the script, compile it and then execute it
        execResult = core.LoadExecScript("script", scriptfile);
        Assert.IsTrue(execResult.Result);

        //--check the content of excel file
        var fileStream = ExcelTestChecker.OpenExcel(PathExcelFilesExec + "datScriptOnExcel1.xlsx");
        Assert.IsNotNull(fileStream);
        var wb = ExcelTestChecker.GetWorkbook(fileStream);

        // r1, c0: 9  -> not modified
        bool res = ExcelTestChecker.CheckCellValue(wb, 0, 1, 0, 9);
        Assert.IsTrue(res);

        // r2, c0: 10 -> modified!
        res = ExcelTestChecker.CheckCellValue(wb, 0, 2, 0, 10);
        Assert.IsTrue(res);

    }

    /// <summary>
    /// The path of the excel file is wrong.
    /// OnExcel instr.
    /// </summary>
    [TestMethod]
    public void OnExcelBasicPathWrong()
    {
        ExecResult execResult;
        LexerowCore core = new LexerowCore();
        string scriptfile = PathScriptFiles + "execOnExcel1Err.lxrw";

        // load the script, compile it and then execute it
        execResult = core.LoadExecScript("script", scriptfile);
        Assert.IsFalse(execResult.Result);
        Assert.AreEqual(ErrorCode.ExecInstrFilePathWrong, execResult.ListError[0].ErrorCode);
    }

}
