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

namespace Lexerow.Core.Tests.CoreLoadLinesExec;

/// <summary>
/// Test the script load, compile and execute from the core.
/// Focus on CreateExcel instruction.
/// Need to have input excel files ready.
/// </summary>
[TestClass]
public class LoadLinesExecCreateExcelTests : BaseTests
{
    /// <summary>
    /// file=CreateExcel("path\file.xlsx")
    /// </summary>
    [TestMethod]
    public void CreateExcelStringOk()
    {
        Result result;
        LexerowCore core = new LexerowCore();

        // create a basic script
        List<string> lines = [
            "file=CreateExcel(" + AddDblQuote(PathExcelFilesExec + "createExcel1.xlsx") +")"
            ];

        // remove the file if it exists
        if (File.Exists(PathExcelFilesExec + "createExcel1.xlsx"))
            File.Delete(PathExcelFilesExec + "createExcel1.xlsx");

        // load the script and compile it
        result = core.LoadLinesScript("script", lines);
        Assert.IsTrue(result.Res);

        result = core.ExecuteScript("script");
        Assert.IsTrue(result.Res);

        // check the file is created
        ExcelFile excelFile = TestExcelChecker.Open(PathExcelFilesExec + "createExcel1.xlsx");
        Assert.IsNotNull(excelFile);

        // check the file has one sheet named xxx
        ExcelSheet excelSheet = TestExcelChecker.GetSheetByName(excelFile, CoreDefinitions.DefaultExcelSheetName);
    }

    /// <summary>
    /// file=CreateExcel("path\file.xlsx", "data")
    /// </summary>
    [TestMethod]
    public void CreateExcelSheetNameOk()
        {
        Result result;
        LexerowCore core = new LexerowCore();
        // create a basic script
        List<string> lines = [
            "file=CreateExcel(" + AddDblQuote(PathExcelFilesExec + "createExcel2.xlsx") + ", \"data\")"
            ];

        // remove the file if it exists
        if (File.Exists(PathExcelFilesExec + "createExcel2.xlsx"))
            File.Delete(PathExcelFilesExec + "createExcel2.xlsx");

        // load the script and compile it
        result = core.LoadLinesScript("script", lines);
        Assert.IsTrue(result.Res);
        result = core.ExecuteScript("script");
        Assert.IsTrue(result.Res);
        // check the file is created
        ExcelFile excelFile = TestExcelChecker.Open(PathExcelFilesExec + "createExcel2.xlsx");
        Assert.IsNotNull(excelFile);
        // check the file has one sheet named data
        ExcelSheet excelSheet = TestExcelChecker.GetSheetByName(excelFile, "data");
    }

    /// <summary>
    // filename is a var, not a string
    /// filename= "path\file.xlsx"
    /// file=CreateExcel(filename)
    /// where
    [TestMethod] 
     public void CreateExcelVarOk()
    {
        Result result;
        LexerowCore core = new LexerowCore();
    
        // create a basic script
        List<string> lines = [
            "filename=" + AddDblQuote(PathExcelFilesExec + "createExcel3.xlsx"),
            "file=CreateExcel(filename)"
        ];

        // remove the file if it exists
        if (File.Exists(PathExcelFilesExec + "createExcel3.xlsx"))
            File.Delete(PathExcelFilesExec + "createExcel3.xlsx");

        // load the script and compile it
        result = core.LoadLinesScript("script", lines);
        Assert.IsTrue(result.Res);
    
        result = core.ExecuteScript("script");
        Assert.IsTrue(result.Res);
    
        // check the file is created
        ExcelFile excelFile = TestExcelChecker.Open(PathExcelFilesExec + "createExcel3.xlsx");
        Assert.IsNotNull(excelFile);

        // check the file has one sheet named default
        ExcelSheet excelSheet = TestExcelChecker.GetSheetByName(excelFile, CoreDefinitions.DefaultExcelSheetName);
    }

}
