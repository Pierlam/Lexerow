using Lexerow.Core.System;
using Lexerow.Core.Tests._20_Utils;
using Lexerow.Core.Tests.Common;
using OpenExcelSdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Tests.CoreLoadFileExec;

/// <summary>
/// Test the script load from txt file, compile and execute from the core.
/// Focus on OnExcel If And/Or instruction.
/// Need to have script and input excel files ready.
/// </summary>
[TestClass]
public class LoadFileExecOnExcelIfAndOrTests : BaseTests
{
    /// <summary>
    /// OnExcel ".\10-ExcelFiles\onExcelIfAndOk.xlsx"
    ///   ForEachRow
    ///     If A.Cell> 10 And A.Cell<20 Then A.Cell=56
    ///   Next
    /// End OnExcel
    /// </summary>
    [TestMethod]
    public void OnExcelIfAndOk()
    {
        LexerowCore core = new LexerowCore();
        string scriptfile = PathScriptFiles + "onExcelIfAndOk.lxrw";

        // load the script, compile it and then execute it
        Result result = core.LoadExecScript("script", scriptfile);
        Assert.IsTrue(result.Res);

        //--check the content of excel file
        ExcelFile excelFile = TestExcelChecker.Open(PathExcelFilesExec + "onExcelIfAndOk.xlsx");
        Assert.IsNotNull(excelFile);

        //--A2: 12 ->56
        bool res = TestExcelChecker.CheckCellValue(excelFile, "A2", 56);
        Assert.IsTrue(res);
    
        //--A3:  8
        res = TestExcelChecker.CheckCellValue(excelFile, "A3", 8);
        Assert.IsTrue(res);

        //--A4:  25
        res = TestExcelChecker.CheckCellValue(excelFile, "A4", 25);
        Assert.IsTrue(res);
    }

}
