using Lexerow.Core.System;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Tests.CoreLoadLinesCompile;

/// <summary>
/// Test the script load from lines and compile from the core.
/// Focus on CopyDataHeader instruction.
/// No need to have input excel file. Check the script content/structure itself.
/// </summary>
[TestClass]
public class LoadLinesCompileCopyHeaderTests
{
    /// <summary>
    /// CopyHeader(string, string)
    /// </summary>
    [TestMethod]
    public void CopyHeaderStringStringOk()
    {
        Result result;
        LexerowCore core = new LexerowCore();

        // create a basic script
        List<string> lines = [
            "file= SelectFiles(\"data.xlsx\")",
            "excelOut= CreateExcel(\"result.xlsx\")",
            "CopyHeader(\"data.xlsx\", \"result.xlsx\")",
            ];

        // load the script and compile it
        result = core.LoadLinesScript("script", lines);
        Assert.IsTrue(result.Res);
    }

    /// <summary>
    /// CopyHeader(excelFileObject, excelFileObject)
    /// </summary>
    [TestMethod]
    public void CopyHeaderExcelFileFileOk()
    {
        Result result;
        LexerowCore core = new LexerowCore();

        // create a basic script
        List<string> lines = [
            "file= SelectFiles(\"data.xlsx\")",
            "excelOut= CreateExcel(\"result.xlsx\")",
            "CopyHeader(file, excelOut)",
            ];

        // load the script and compile it
        result = core.LoadLinesScript("script", lines);
        Assert.IsTrue(result.Res);
    }

    /// <summary>
    /// CopyHeader(excelFileObject, excelFileObject)
    /// </summary>
    [TestMethod]
    public void CopyHeaderExcelFileListOfExcelFileObjectOk()
    {
        Result result;
        LexerowCore core = new LexerowCore();

        // create a basic script
        List<string> lines = [
            "file= SelectFiles(\"*.xlsx\")",
            "excelOut= CreateExcel(\"result.xlsx\")",
            "CopyHeader(file, excelOut)",
            ];

        // load the script and compile it
        result = core.LoadLinesScript("script", lines);
        Assert.IsTrue(result.Res);
    }

    /// <summary>
    /// var= CopyHeader(string, string)
    /// </summary>
    [TestMethod]
    public void ResCopyHeaderStringStringWrong()
    {
        Result result;
        LexerowCore core = new LexerowCore();

        // create a basic script
        List<string> lines = [
            "file= SelectFiles(\"data.xlsx\")",
            "excelOut= CreateExcel(\"result.xlsx\")",
            "var= CopyHeader(\"data.xlsx\", \"result.xlsx\")",
            ];

        // load the script and compile it
        result = core.LoadLinesScript("script", lines);
        Assert.IsFalse(result.Res);
        Assert.AreEqual(1, result.ListError.Count); 
        Assert.AreEqual(ErrorCode.ParserVarWrongRightPart, result.ListError[0].ErrorCode);
    }

    /// <summary>
    /// var= CopyHeader(string)
    /// </summary>
    [TestMethod]
    public void ResCopyHeaderStringWrong()
    {
        Result result;
        LexerowCore core = new LexerowCore();

        // create a basic script
        List<string> lines = [
            "CopyHeader(\"data.xlsx\")",
            ];

        // load the script and compile it
        result = core.LoadLinesScript("script", lines);
        Assert.IsFalse(result.Res);
        Assert.AreEqual(1, result.ListError.Count);
        Assert.AreEqual(ErrorCode.ParserFctParamCountWrong, result.ListError[0].ErrorCode);
    }

}
