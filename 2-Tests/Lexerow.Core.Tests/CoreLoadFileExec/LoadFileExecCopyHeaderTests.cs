using DocumentFormat.OpenXml.Spreadsheet;
using Lexerow.Core.System;
using Lexerow.Core.System.GenDef;
using Lexerow.Core.Tests._20_Utils;
using Lexerow.Core.Tests.Common;
using OpenExcelSdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Tests.CoreLoadFileExec;

[TestClass]
public class LoadFileExecCopyHeaderTests : BaseTests
{
    /// <summary>
    /// file=SelectFiles("path\copyHeader1.xlsx")    
    /// fileRes=CreateExcel("path\copyHeaderRes1.xlsx")
    /// CopyHeader(file, fileRes)
    /// </summary>
    [TestMethod]
    public void CopyHeaderFileFileResOk()
    {

        // remove the file if it exists
        if (File.Exists(PathExcelFilesExec + "copyHeaderRes1.xlsx"))
            File.Delete(PathExcelFilesExec + "copyHeaderRes1.xlsx");

        Result result;
        LexerowCore core = new LexerowCore();
        string scriptfile = PathScriptFiles + "copyHeader1.lxrw";

        // load the script, compile it and then execute it
        result = core.LoadExecScript("script", scriptfile);
        Assert.IsTrue(result.Res);


        //--check the file is created
        ExcelFile excelFile = TestExcelChecker.Open(PathExcelFilesExec + "copyHeaderRes1.xlsx");
        Assert.IsNotNull(excelFile);

        // check the file has one sheet named xxx
        ExcelSheet excelSheet = TestExcelChecker.GetSheetByName(excelFile, CoreDefinitions.DefaultExcelSheetName);

        // check cells 
        bool res = TestExcelChecker.CheckCellValue(excelFile, "A1", "Id");
        Assert.IsTrue(res);
        res = TestExcelChecker.CheckCellValue(excelFile, "B1", "lastname");
        Assert.IsTrue(res);
        res = TestExcelChecker.CheckCellValue(excelFile, "C1", "firstname");
        Assert.IsTrue(res);
        res = TestExcelChecker.CheckCellValue(excelFile, "D1", "cost");
        Assert.IsTrue(res);
        res= TestExcelChecker.CheckCellNull(excelFile, "E1");
        Assert.IsTrue(res);
    }

    /// <summary>
    /// # "path\copyHeaderRes2.xlsx"  exists and be closed
    /// 
    /// CopyHeader("path\copyHeader2.xlsx", "path\copyHeaderRes2.xlsx")
    /// </summary>
    [TestMethod]
    public void CopyHeaderStringStringOk()
    {
        Result result;
        LexerowCore core = new LexerowCore();
        string scriptfile = PathScriptFiles + "copyHeader2.lxrw";

        // load the script, compile it and then execute it
        result = core.LoadExecScript("script", scriptfile);
        Assert.IsTrue(result.Res);


        //--check the file is created
        ExcelFile excelFile = TestExcelChecker.Open(PathExcelFilesExec + "copyHeaderRes2.xlsx");
        Assert.IsNotNull(excelFile);

        // check the file has one sheet named xxx
        ExcelSheet excelSheet = TestExcelChecker.GetSheetByName(excelFile, CoreDefinitions.DefaultExcelSheetName);

        // check cells 
        bool res = TestExcelChecker.CheckCellValue(excelFile, "A1", "Id");
        Assert.IsTrue(res);
        res = TestExcelChecker.CheckCellValue(excelFile, "B1", "lastname");
        Assert.IsTrue(res);
        res = TestExcelChecker.CheckCellValue(excelFile, "C1", "firstname");
        Assert.IsTrue(res);
        res = TestExcelChecker.CheckCellValue(excelFile, "D1", "cost");
        Assert.IsTrue(res);
        res = TestExcelChecker.CheckCellNull(excelFile, "E1");
        Assert.IsTrue(res);
    }


    /// <summary>
    /// fileRes=CreateExcel("path\copyHeaderResErr1.xlsx")
    /// 
    /// # error! the file is already opened! by the previous instr CreateExcel
    /// CopyHeader("path\copyHeader2.xlsx", "path\copyHeaderResErr1.xlsx")
    /// </summary>
    [TestMethod]
    public void CopyHeaderStringStringErr1()
    {
        // remove the file if it exists
        if (File.Exists(PathExcelFilesExec + "copyHeaderResErr1.xlsx"))
            File.Delete(PathExcelFilesExec + "copyHeaderResErr1.xlsx");

        Result result;
        LexerowCore core = new LexerowCore();
        string scriptfile = PathScriptFiles + "copyHeaderErr1.lxrw";

        // load the script, compile it and then execute it
        result = core.LoadExecScript("script", scriptfile);
        Assert.IsFalse(result.Res);
        Assert.AreEqual(1, result.ListError.Count);
        Assert.AreEqual(ErrorCode.ExecUnableCopyHeader, result.ListError[0].ErrorCode);
    }

    /// <summary>
    /// Both source and taget excel file does not exist.
    /// </summary>
    [TestMethod]
    public void CopyHeaderStringStringErr2()
    {
        Result result;
        LexerowCore core = new LexerowCore();
        string scriptfile = PathScriptFiles + "copyHeaderErr2.lxrw";

        // load the script, compile it and then execute it
        result = core.LoadExecScript("script", scriptfile);
        Assert.IsFalse(result.Res);
        Assert.AreEqual(1, result.ListError.Count);
        Assert.AreEqual(ErrorCode.ExecUnableCopyHeader, result.ListError[0].ErrorCode);
    }

    /// <summary>
    /// source sheet is empty.
    /// </summary>
    [TestMethod]
    public void CopyHeaderStringStringErr3()
    {
        Result result;
        LexerowCore core = new LexerowCore();
        string scriptfile = PathScriptFiles + "copyHeaderErr3.lxrw";

        // load the script, compile it and then execute it
        result = core.LoadExecScript("script", scriptfile);
        Assert.IsFalse(result.Res);
        Assert.AreEqual(1, result.ListError.Count);
        Assert.AreEqual(ErrorCode.ExecExcelSheetEmpty, result.ListError[0].ErrorCode);
    }

    /// <summary>
    /// Target sheet is empty.
    /// </summary>
    [TestMethod]
    public void CopyHeaderStringStringErr4()
    {
        Result result;
        LexerowCore core = new LexerowCore();
        string scriptfile = PathScriptFiles + "copyHeaderErr4.lxrw";

        // load the script, compile it and then execute it
        result = core.LoadExecScript("script", scriptfile);
        Assert.IsFalse(result.Res);
        Assert.AreEqual(1, result.ListError.Count);
        Assert.AreEqual(ErrorCode.ExecExcelSheetNotEmpty, result.ListError[0].ErrorCode);

        Assert.IsTrue(result.ListError[0].Param2.EndsWith("copyHeaderResErr4.xlsx"));
    }
}
