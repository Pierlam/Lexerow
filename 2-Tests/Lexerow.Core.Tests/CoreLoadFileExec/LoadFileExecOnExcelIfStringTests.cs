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
public class LoadFileExecOnExcelIfStringTests : BaseTests
{
    /// <summary>
    /// OnExcel ".\10-ExcelFiles\IfACellEqString.xlsx"
    ///   ForEachRow
    ///     If A.Cell="Hello" Then A.Cell="Bonjour"
    ///   Next
    ///   End OnExcel
    /// </summary>
    [TestMethod]
    public void IfACellEqStringOk()
    {
        LexerowCore core = new LexerowCore();
        string scriptfile = PathScriptFiles + "IfACellEqString.lxrw";

        // load the script, compile it and then execute it
        Result result = core.LoadExecScript("script", scriptfile);
        Assert.IsTrue(result.Res);

        //--check the content of excel file
        ExcelFile excelFile = TestExcelChecker.Open(PathExcelFilesExec + "IfACellEqString.xlsx");
        Assert.IsNotNull(excelFile);

        // A2: Hello -> Bonjour
        bool res = TestExcelChecker.CheckCellValue(excelFile, "A2", "Bonjour");
        Assert.IsTrue(res);

        // A3: HELLO -> Bonjour, case Insensitive by default!
        res = TestExcelChecker.CheckCellValue(excelFile, "A3", "Bonjour");
        Assert.IsTrue(res);

        // A4: salute 
        res = TestExcelChecker.CheckCellValue(excelFile, "A4", "salute");
        Assert.IsTrue(res);

        // A5: Ave 
        res = TestExcelChecker.CheckCellValue(excelFile, "A5", "Ave");
        Assert.IsTrue(res);
    }

    /// <summary>
    /// OnExcel ".\10-ExcelFiles\IfACellEqString.xlsx"
    ///   ForEachRow
    ///     If A.Cell<>"Hello" Then A.Cell="Bonjour"
    ///   Next
    ///   End OnExcel
    /// </summary>
    [TestMethod]
    public void IfACellNotEqStringOk()
    {
        LexerowCore core = new LexerowCore();
        string scriptfile = PathScriptFiles + "IfACellNotEqString.lxrw";

        // load the script, compile it and then execute it
        Result result = core.LoadExecScript("script", scriptfile);
        Assert.IsTrue(result.Res);

        //--check the content of excel file
        ExcelFile excelFile = TestExcelChecker.Open(PathExcelFilesExec + "IfACellNotEqString.xlsx");
        Assert.IsNotNull(excelFile);

        // A2: Hello
        bool res = TestExcelChecker.CheckCellValue(excelFile, "A2", "Hello");
        Assert.IsTrue(res);

        // A3: HELLO, case Insensitive by default!
        res = TestExcelChecker.CheckCellValue(excelFile, "A3", "HELLO");
        Assert.IsTrue(res);

        // A4: salute  -> Bonjour
        res = TestExcelChecker.CheckCellValue(excelFile, "A4", "Bonjour");
        Assert.IsTrue(res);

        // A5: Ave -> Bonjour
        res = TestExcelChecker.CheckCellValue(excelFile, "A5", "Bonjour");
        Assert.IsTrue(res);
    }

    /// <summary>
    /// $StrCompareCaseSensitive=true
    /// OnExcel ".\10-ExcelFiles\IfACellEqString.xlsx"
    ///   ForEachRow
    ///     If A.Cell="Hello" Then A.Cell="Bonjour"
    ///   Next
    ///   End OnExcel
    /// </summary>
    [TestMethod]
    public void IfACellEqStringCaseSensitiveOk()
    {
        LexerowCore core = new LexerowCore();
        string scriptfile = PathScriptFiles + "IfACellEqStringCaseSensitive.lxrw";

        // load the script, compile it and then execute it
        Result result = core.LoadExecScript("script", scriptfile);
        Assert.IsTrue(result.Res);

        //--check the content of excel file
        ExcelFile excelFile = TestExcelChecker.Open(PathExcelFilesExec + "IfACellEqStringCaseSensitive.xlsx");
        Assert.IsNotNull(excelFile);

        // A2: Hello -> Bonjour
        bool res = TestExcelChecker.CheckCellValue(excelFile, "A2", "Bonjour");
        Assert.IsTrue(res);

        // A3: HELLO stay!, case sensitive=true
        res = TestExcelChecker.CheckCellValue(excelFile, "A3", "HELLO");
        Assert.IsTrue(res);

        // A4: salute 
        res = TestExcelChecker.CheckCellValue(excelFile, "A4", "salute");
        Assert.IsTrue(res);

        // A5: Ave 
        res = TestExcelChecker.CheckCellValue(excelFile, "A5", "Ave");
        Assert.IsTrue(res);
    }

}
