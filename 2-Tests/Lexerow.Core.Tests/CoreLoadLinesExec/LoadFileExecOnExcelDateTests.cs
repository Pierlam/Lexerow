using Lexerow.Core.System;
using Lexerow.Core.Tests._20_Utils;
using Lexerow.Core.Tests.Common;
using OpenExcelSdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Tests.CoreLoadLinesExec;


/// <summary>
/// Test the script load from txt file, compile and execute from the core.
/// Focus on OnExcel and Date instruction.
/// Need to have script and input excel files ready.
/// </summary>
[TestClass]
public class LoadFileExecOnExcelDateTests : BaseTests
{
    /// <summary>
    /// If A.Cell <= Date(2023,11,14) Then B.Cell="yes"
    /// </summary>
    [TestMethod]
    public void OnExcelIfDateThenYesOk()
    {
        Result result;
        LexerowCore core = new LexerowCore();
        string scriptfile = PathScriptFiles + "OnExcelIfDateThenYesOk.lxrw";

        // load the script, compile it and then execute it
        result = core.LoadExecScript("script", scriptfile);
        Assert.IsTrue(result.Res);

        //--check the content of excel file
        ExcelFile excelFile = TestExcelChecker.Open(PathExcelFilesExec + "OnExcelIfDateThenYesOk.xlsx");
        Assert.IsNotNull(excelFile);

        //==>check some cell value

        bool res = TestExcelChecker.CheckCellValue(excelFile, "B2", "yes");
        Assert.IsTrue(res);

        res = TestExcelChecker.CheckCellValue(excelFile, "B3", "no");
        Assert.IsTrue(res);

        // A4: 03:40:25
        res = TestExcelChecker.CheckCellValue(excelFile, "B4", "no");
        Assert.IsTrue(res);

        // A5: special case!! 03/05/2018  12:30:45   -> compare DateTime with DateOnly -> ok
        res = TestExcelChecker.CheckCellValue(excelFile, "B5", "yes");
        Assert.IsTrue(res);
    }

    /// <summary>
    /// If A.Cell <= 10 Then B.Cell=Date(2025, 11, 30)
    /// </summary>
    [TestMethod]
    public void OnExcelIfThenDateOk()
    {
        Result result;
        LexerowCore core = new LexerowCore();
        string scriptfile = PathScriptFiles + "OnExcelIfThenDateOk.lxrw";

        // load the script, compile it and then execute it
        result = core.LoadExecScript("script", scriptfile);
        Assert.IsTrue(result.Res);

        //--check the content of excel file
        ExcelFile excelFile = TestExcelChecker.Open(PathExcelFilesExec + "OnExcelIfThenDateOk.xlsx");
        Assert.IsNotNull(excelFile);

        //==>check some cell value
        bool res;

        // NOT modified
        res = TestExcelChecker.CheckCellValue(excelFile, "A2", 10);
        Assert.IsTrue(res);

        res = TestExcelChecker.CheckCellValue(excelFile, "A3", 12);
        Assert.IsTrue(res);

        // only one cell updated
        res = TestExcelChecker.CheckCellValue(excelFile, "B2", new DateOnly(2025, 11, 30));
        Assert.IsTrue(res);

        // TODO: keep it??
        res = TestExcelChecker.CheckCellStyleNumberFormatId(excelFile, "B2", 14);
        Assert.IsTrue(res);

        res = TestExcelChecker.CheckCellValue(excelFile, "B3", "greater");
        Assert.IsTrue(res);

        res = TestExcelChecker.CheckCellValue(excelFile, "B4", "hour");
        Assert.IsTrue(res);

        res = TestExcelChecker.CheckCellValue(excelFile, "B5", "dateTime");
        Assert.IsTrue(res);
    }

    /// <summary>
    /// If A.Cell <= Date(2023,11,14) Then A.Cell=Date(2025,3, 12)
    /// Apply the default format, defined in the var $DateFormat.
    /// </summary>
    [TestMethod]
    public void OnExcelIfDateThenDateOk()
    {
        Result result;
        LexerowCore core = new LexerowCore();
        string scriptfile = PathScriptFiles + "OnExcelIfDateThenDateOk.lxrw";

        // load the script, compile it and then execute it
        result = core.LoadExecScript("script", scriptfile);
        Assert.IsTrue(result.Res);

        //--check the content of excel file
        ExcelFile excelFile = TestExcelChecker.Open(PathExcelFilesExec + "OnExcelIfDateThenDateOk.xlsx");
        Assert.IsNotNull(excelFile);

        //==>check some cell value

        // A2: 10/02/2019 -> 12/03/2025
        bool res = TestExcelChecker.CheckCellValue(excelFile, "A2", new DateOnly(2025, 03, 12));
        Assert.IsTrue(res);

        // A3: 03/05/2025
        res = TestExcelChecker.CheckCellValue(excelFile, "A3", new DateOnly(2025, 05, 03));
        Assert.IsTrue(res);

        // A4: 03:40:25
        res = TestExcelChecker.CheckCellValue(excelFile, "A4", new TimeOnly(03, 40, 25));
        Assert.IsTrue(res);

        // A5: 03/05/2018  12:30:45 -> 12/03/2025
        res = TestExcelChecker.CheckCellValue(excelFile, "A5", new DateOnly(2025, 03, 12));
        Assert.IsTrue(res);
    }

    /// <summary>
    /// a=Date(2019,11,14)
    /// If A.Cell <= a Then A.Cell=Date(2025,12, 30)
    /// Apply the default format, defined in the var $DateFormat.
    /// </summary>
    [TestMethod]
    public void OnExcelIfVarDateThenDateOk()
    {
        Result result;
        bool res;

        LexerowCore core = new LexerowCore();
        string scriptfile = PathScriptFiles + "OnExcelIfVarDateThenDate.lxrw";

        // load the script, compile it and then execute it
        result = core.LoadExecScript("script", scriptfile);
        Assert.IsTrue(result.Res);

        //--check the content of excel file
        ExcelFile excelFile = TestExcelChecker.Open(PathExcelFilesExec + "OnExcelIfVarDateThenDate.xlsx");
        Assert.IsNotNull(excelFile);

        //==>check some cell value

        // A2: 14/11/2019 -> 30/12/2025
        res = TestExcelChecker.CheckCellValue(excelFile, "A2", new DateOnly(2025, 12, 30));
        Assert.IsTrue(res);
        // built-in numberFormatId expected: 14: dd/mm/yyyy
        res = TestExcelChecker.CheckCellStyleNumberFormatId(excelFile, "A2", 14);
        Assert.IsTrue(res);

        // A3: 10/05/2020 
        res = TestExcelChecker.CheckCellValue(excelFile, "A3", new DateOnly(2020, 05, 10));
        Assert.IsTrue(res);

        // A4: 25/09/1987 -> 30/12/2025
        res = TestExcelChecker.CheckCellValue(excelFile, "A4", new DateOnly(2025, 12, 30));
        Assert.IsTrue(res);

        // A5: hello
        res = TestExcelChecker.CheckCellValue(excelFile, "A5", "hello");
        Assert.IsTrue(res);
    }

    /// <summary>
    /// a=Date(2025,12, 30)
    /// If A.Cell <= Date(2019,11,14) Then A.Cell=a
    /// Apply the default format, defined in the var $DateFormat.
    /// </summary>
    [TestMethod]
    public void OnExcelIfDateThenVarDateOk()
    {
        Result result;
        bool res;

        LexerowCore core = new LexerowCore();
        string scriptfile = PathScriptFiles + "OnExcelIfDateThenVarDate.lxrw";

        // load the script, compile it and then execute it
        result = core.LoadExecScript("script", scriptfile);
        Assert.IsTrue(result.Res);

        //--check the content of excel file
        ExcelFile excelFile = TestExcelChecker.Open(PathExcelFilesExec + "OnExcelIfDateThenVarDate.xlsx");
        Assert.IsNotNull(excelFile);

        //==>check some cell value

        // A2: 14/11/2019 -> 30/12/2025
        res = TestExcelChecker.CheckCellValue(excelFile, "A2", new DateOnly(2025, 12, 30));
        Assert.IsTrue(res);
        // built-in numberFormatId expected: 14: dd/mm/yyyy
        res = TestExcelChecker.CheckCellStyleNumberFormatId(excelFile, "A2", 14);
        Assert.IsTrue(res);

        // A3: 10/05/2020 
        res = TestExcelChecker.CheckCellValue(excelFile, "A3", new DateOnly(2020, 05, 10));
        Assert.IsTrue(res);

        // A4: 25/09/1987 -> 30/12/2025
        res = TestExcelChecker.CheckCellValue(excelFile, "A4", new DateOnly(2025, 12, 30));
        Assert.IsTrue(res);

        // A5: hello
        res = TestExcelChecker.CheckCellValue(excelFile, "A5", "hello");
        Assert.IsTrue(res);
    }

    /// <summary>
    /// year=2025
    /// month=12
    /// day=30
    /// If A.Cell <= Date(2019,11,14) Then A.Cell=Date(year,month,day)
    /// Apply the default format, defined in the var $DateFormat.
    /// </summary>
    [TestMethod]
    public void OnExcelIfDateThenDate3VarsOk()
    {
        Result result;
        bool res;

        LexerowCore core = new LexerowCore();
        string scriptfile = PathScriptFiles + "OnExcelIfDateThenDate3Vars.lxrw";

        // load the script, compile it and then execute it
        result = core.LoadExecScript("script", scriptfile);
        Assert.IsTrue(result.Res);

        //--check the content of excel file
        ExcelFile excelFile = TestExcelChecker.Open(PathExcelFilesExec + "OnExcelIfDateThenDate3Vars.xlsx");
        Assert.IsNotNull(excelFile);

        //==>check some cell value

        // A2: 14/11/2019 -> 30/12/2025
        res = TestExcelChecker.CheckCellValue(excelFile, "A2", new DateOnly(2025, 12, 30));
        Assert.IsTrue(res);
        // built-in numberFormatId expected: 14: dd/mm/yyyy
        res = TestExcelChecker.CheckCellStyleNumberFormatId(excelFile, "A2", 14);
        Assert.IsTrue(res);

        // A3: 10/05/2020 
        res = TestExcelChecker.CheckCellValue(excelFile, "A3", new DateOnly(2020, 05, 10));
        Assert.IsTrue(res);

        // A4: 25/09/1987 -> 30/12/2025
        res = TestExcelChecker.CheckCellValue(excelFile, "A4", new DateOnly(2025, 12, 30));
        Assert.IsTrue(res);
        // built-in numberFormatId expected: 14: dd/mm/yyyy
        res = TestExcelChecker.CheckCellStyleNumberFormatId(excelFile, "A4", 14);
        Assert.IsTrue(res);

        // A5: hello
        res = TestExcelChecker.CheckCellValue(excelFile, "A5", "hello");
        Assert.IsTrue(res);
    }


    /// <summary>
    /// $DateFormat="yyyy-mm-dd"
    /// If A.Cell <= Date(2019,11,14) Then A.Cell=Date(2025,12, 30)
    /// Apply the default format, defined in the var $DateFormat.
    /// </summary>
    [TestMethod]
    public void OnExcelIfDateThenDateFormatOk()
    {
        Result result;
        bool res;

        LexerowCore core = new LexerowCore();
        string scriptfile = PathScriptFiles + "OnExcelIfDateThenDateFormat.lxrw";

        // load the script, compile it and then execute it
        result = core.LoadExecScript("script", scriptfile);
        Assert.IsTrue(result.Res);

        //--check the content of excel file
        ExcelFile excelFile = TestExcelChecker.Open(PathExcelFilesExec + "OnExcelIfDateThenDateFormat.xlsx");
        Assert.IsNotNull(excelFile);

        //==>check some cell value

        // A2: 14/11/2019 -> 30/12/2025
        res = TestExcelChecker.CheckCellValue(excelFile, "A2", new DateOnly(2025, 12, 30));
        Assert.IsTrue(res);
        // built-in numberFormatId expected: 14: dd/mm/yyyy
        res = TestExcelChecker.CheckCellStyleNumberFormat(excelFile, "A2", "yyyy-mm-dd");
        Assert.IsTrue(res);

        // A3: 10/05/2020 
        res = TestExcelChecker.CheckCellValue(excelFile, "A3", new DateOnly(2020, 05, 10));
        Assert.IsTrue(res);

        // A4: 25/09/1987 -> 30/12/2025
        res = TestExcelChecker.CheckCellValue(excelFile, "A4", new DateOnly(2025, 12, 30));
        Assert.IsTrue(res);
        res = TestExcelChecker.CheckCellStyleNumberFormat(excelFile, "A2", "yyyy-mm-dd");
        Assert.IsTrue(res);

        // A5: hello
        res = TestExcelChecker.CheckCellValue(excelFile, "A5", "hello");
        Assert.IsTrue(res);
    }

}