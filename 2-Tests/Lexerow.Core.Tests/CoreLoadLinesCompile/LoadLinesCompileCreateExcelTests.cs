using Lexerow.Core.System;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Tests.CoreLoadLinesCompile;

/// <summary>
/// Test the script load from lines and compile from the core.
/// Focus on CreateExcel instruction.
/// No need to have input excel file. Check the script content/structure itself.
/// </summary>
[TestClass]
public class LoadLinesCompileCreateExcelTests
{
    /// <summary>
    /// file= CreateExcel("mydata.xlsx")
    /// </summary>
    [TestMethod]
    public void CreateExcelStringOk()
    {
        Result result;
        LexerowCore core = new LexerowCore();

        // create a basic script
        List<string> lines = [
            "file= CreateExcel(\"mydata.xlsx\")",
            ];

        // load the script and compile it
        result = core.LoadLinesScript("script", lines);
        Assert.IsTrue(result.Res);
    }


    /// <summary>
    /// filename= "mydata.xlsx"
    ///  file= CreateExcel(filename)
    /// </summary>
    [TestMethod]
    public void CreateExcelVarOk()
    {
        Result result;
        LexerowCore core = new LexerowCore();

        // create a basic script
        List<string> lines = [
            "filename= \"mydata.xlsx\"",
            "file= CreateExcel(filename)",
            ];

        // load the script and compile it
        result = core.LoadLinesScript("script", lines);
        Assert.IsTrue(result.Res);
    }

    /// <summary>
    /// file= CreateExcel("mydata.xlsx", "sheet1")
    /// </summary>
    [TestMethod]
    public void CreateExcelStringSheetOk()
    {
        Result result;
        LexerowCore core = new LexerowCore();

        // create a basic script
        List<string> lines = [
            "file= CreateExcel(\"mydata.xlsx\", \"sheet1\")",
            ];

        // load the script and compile it
        result = core.LoadLinesScript("script", lines);
        Assert.IsTrue(result.Res);
    }


    /// <summary>
    /// filename= "mydata.xlsx"
    ///  file= CreateExcel(filename, "sheet1")
    /// </summary>
    [TestMethod]
    public void CreateExcelVarSheetOk()
    {
        Result result;
        LexerowCore core = new LexerowCore();

        // create a basic script
        List<string> lines = [
            "filename=\"mydata.xlsx\"",
            "file= CreateExcel(filename,\"sheet1\")",
            ];

        // load the script and compile it
        result = core.LoadLinesScript("script", lines);
        Assert.IsTrue(result.Res);
    }

    /// <summary>
    ///  file= CreateExcel()
    /// </summary>
    [TestMethod]
    public void CreateExcelNoParamWrong()
    {
        Result result;
        LexerowCore core = new LexerowCore();

        // create a basic script
        List<string> lines = [
            "file= CreateExcel()",
            ];

        // load the script and compile it
        result = core.LoadLinesScript("script", lines);
        Assert.IsFalse(result.Res);
    }

    /// <summary>
    ///  CreateExcel()
    /// </summary>
    [TestMethod]
    public void CreateExcelNoReturnWrong()
    {
        Result result;
        LexerowCore core = new LexerowCore();

        // create a basic script
        List<string> lines = [
            "CreateExcel(\"mydata.xlsx\")",
            ];

        // load the script and compile it
        result = core.LoadLinesScript("script", lines);
        Assert.IsFalse(result.Res);
    }

    /// <summary>
    ///  file=CreateExcel(12)
    /// </summary>
    [TestMethod]
    public void CreateExcelParamIntWrong()
    {
        Result result;
        LexerowCore core = new LexerowCore();

        // create a basic script
        List<string> lines = [
            "file=CreateExcel(12)",
            ];

        // load the script and compile it
        result = core.LoadLinesScript("script", lines);
        Assert.IsFalse(result.Res);
    }

}
