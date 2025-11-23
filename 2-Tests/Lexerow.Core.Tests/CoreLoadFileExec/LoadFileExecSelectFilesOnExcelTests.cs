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
/// Focus on SelectFiles + OnExcel instruction.
/// Need to have script and input excel files ready.
/// </summary>
[TestClass]
public class LoadFileExecSelectFilesOnExcelTests : BaseTests
{
    /// <summary>
    /// Test SelectFiles + OnExcel
    /// </summary>
    [TestMethod]
    public void SelectFilesOnExcelBasicOk()
    {
        Result result;
        LexerowCore core = new LexerowCore();
        string scriptfile = PathScriptFiles + "execSelectFilesOnExcel.lxrw";

        // load the script, compile it and then execute it
        result = core.LoadExecScript("script", scriptfile);
        Assert.IsTrue(result.Res);

        //--check result insights
        Assert.AreEqual(1, result.Insights.FileTotalCount);
        Assert.AreEqual(1, result.Insights.SheetTotalCount);
        Assert.AreEqual(2, result.Insights.RowTotalCount);
        Assert.AreEqual(1, result.Insights.IfCondMatchTotalCount);

        //--check the content of excel file
        var fileStream = TestExcelChecker.OpenExcel(PathExcelFilesExec + "execSelectFilesOnExcel.xlsx");
        Assert.IsNotNull(fileStream);
        var wb = TestExcelChecker.GetWorkbook(fileStream);

        // r1, c0: 9  -> not modified
        bool res = TestExcelChecker.CheckCellValue(wb, 0, 1, 0, 9);
        Assert.IsTrue(res);

        // r2, c0: 10 -> modified!
        res = TestExcelChecker.CheckCellValue(wb, 0, 2, 0, 10);
        Assert.IsTrue(res);
    }

}
