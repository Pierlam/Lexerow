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
/// Test OnExcel If String condition cases.
/// </summary>
[TestClass]
public class LoadFileExecOnExcelFirstRowTests : BaseTests
{
    [TestMethod]
    public void OnExcelFirstRowValueOk()
    {
        LexerowCore core = new LexerowCore();
        string scriptfile = PathScriptFiles + "onExcelFirstRowValue.lxrw";

        // load the script, compile it and then execute it
        Result result = core.LoadExecScript("script", scriptfile);
        Assert.IsTrue(result.Res);

        //--check the content of excel file
        ExcelFile excelFile = TestExcelChecker.Open(PathExcelFilesExec + "onExcelFirstRowValue.xlsx");
        Assert.IsNotNull(excelFile);

        //==>check some cell value

        bool res = TestExcelChecker.CheckCellValue(excelFile, "A2", 9);
        Assert.IsTrue(res);

        res = TestExcelChecker.CheckCellValue(excelFile, "A3", 12);
        Assert.IsTrue(res);

        res = TestExcelChecker.CheckCellValue(excelFile, "A4", 27);
        Assert.IsTrue(res);
    }

    [TestMethod]
    public void OnExcelFirstRowVarOk()
    {
        Result result;
        LexerowCore core = new LexerowCore();
        string scriptfile = PathScriptFiles + "onExcelFirstRowVar.lxrw";

        // load the script, compile it and then execute it
        result = core.LoadExecScript("script", scriptfile);
        Assert.IsTrue(result.Res);

        //--check the content of excel file
        ExcelFile excelFile = TestExcelChecker.Open(PathExcelFilesExec + "onExcelFirstRowVar.xlsx");
        Assert.IsNotNull(excelFile);

        //==>check some cell value

        bool res = TestExcelChecker.CheckCellValue(excelFile, "A2", 9);
        Assert.IsTrue(res);

        res = TestExcelChecker.CheckCellValue(excelFile, "A3", 12);
        Assert.IsTrue(res);

        res = TestExcelChecker.CheckCellValue(excelFile, "A4", 27);
        Assert.IsTrue(res);
    }

    [TestMethod]
    public void OnExcelFirstRowVarVarOk()
    {
        Result result;
        LexerowCore core = new LexerowCore();
        string scriptfile = PathScriptFiles + "onExcelFirstRowVarVar.lxrw";

        //=> load the script, compile it and then execute it
        result = core.LoadExecScript("script", scriptfile);
        Assert.IsTrue(result.Res);

        //=> check the content of excel file
        ExcelFile excelFile = TestExcelChecker.Open(PathExcelFilesExec + "onExcelFirstRowVarVar.xlsx");
        Assert.IsNotNull(excelFile);

        //==>check some cell value

        bool res = TestExcelChecker.CheckCellValue(excelFile, "A2", 9);
        Assert.IsTrue(res);

        res = TestExcelChecker.CheckCellValue(excelFile, "A3", 12);
        Assert.IsTrue(res);

        res = TestExcelChecker.CheckCellValue(excelFile, "A4", 27);
        Assert.IsTrue(res);
    }

}
