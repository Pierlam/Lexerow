using Lexerow.Core.System;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Tests.CoreLoadLinesCompile;

[TestClass]
public class LoadLinesCompileCopyRowTests
{
    /// <summary>
    /// Basic version.
    /// CopyRow(string)
    /// </summary>
    [TestMethod]
    public void CopyRowOk()
    {
        Result result;
        LexerowCore core = new LexerowCore();

        // create a basic script
        List<string> lines = [
            "file= SelectFiles(\"data.xlsx\")",
            "fileRes= CreateExcel(\"result.xlsx\")",
            "CopyHeader(\"data.xlsx\", \"result.xlsx\")",
            "OnExcel file",
            "  ForEach Row",
            "    If A.Cell >10 Then",
            "      CopyRow(fileRes)",
            "    End if",
            "  Next",
            "End OnExcel"
            ];

        // load the script and compile it
        result = core.LoadLinesScript("script", lines);
        Assert.IsTrue(result.Res);
    }

    /// <summary>
    /// file var does not exists.
    /// CopyRow(string)
    /// </summary>
    [TestMethod]
    public void CopyRowFileDoesnotExistsWrong()
    {
        Result result;
        LexerowCore core = new LexerowCore();

        // create a basic script
        List<string> lines = [
            "file= SelectFiles(\"data.xlsx\")",
            "fileRes= CreateExcel(\"result.xlsx\")",
            "CopyHeader(\"data.xlsx\", \"result.xlsx\")",
            "OnExcel file",
            "  ForEach Row",
            "    If A.Cell >10 Then",
            "      CopyRow(fileWrong)",
            "    End if",
            "  Next",
            "End OnExcel"
            ];

        // load the script and compile it
        result = core.LoadLinesScript("script", lines);
        Assert.IsFalse(result.Res);
        Assert.AreEqual(ErrorCode.ParserVarOrFctNameNotDefined, result.ListError[0].ErrorCode);
    }

    /// <summary>
    /// file var does not exists.
    /// CopyRow(string)
    /// </summary>
    [TestMethod]
    public void CopyRowFile12Wrong()
    {
        Result result;
        LexerowCore core = new LexerowCore();

        // create a basic script
        List<string> lines = [
            "file= SelectFiles(\"data.xlsx\")",
            "fileRes= CreateExcel(\"result.xlsx\")",
            "CopyHeader(\"data.xlsx\", \"result.xlsx\")",
            "OnExcel file",
            "  ForEach Row",
            "    If A.Cell >10 Then",
            "      CopyRow(12)",
            "    End if",
            "  Next",
            "End OnExcel"
            ];

        // load the script and compile it
        result = core.LoadLinesScript("script", lines);
        Assert.IsFalse(result.Res);
        Assert.AreEqual(ErrorCode.ParserFctParamTypeWrong, result.ListError[0].ErrorCode);
    }

    /// <summary>
    /// wrong param type, must be an excel file object
    /// CopyRow(string)
    /// </summary>
    [TestMethod]
    public void CopyRowFileStringWrong()
    {
        Result result;
        LexerowCore core = new LexerowCore();

        // create a basic script
        List<string> lines = [
            "file= SelectFiles(\"data.xlsx\")",
            "fileRes= CreateExcel(\"result.xlsx\")",
            "CopyHeader(\"data.xlsx\", \"result.xlsx\")",
            "OnExcel file",
            "  ForEach Row",
            "    If A.Cell >10 Then",
            "      CopyRow(\"data.xlsx\")",
            "    End if",
            "  Next",
            "End OnExcel"
            ];

        // load the script and compile it
        result = core.LoadLinesScript("script", lines);
        Assert.IsFalse(result.Res);
        Assert.AreEqual(ErrorCode.ParserFctParamTypeWrong, result.ListError[0].ErrorCode);
    }

    /// <summary>
    /// too many parameters.
    /// CopyRow(filesRes, string)
    /// </summary>
    [TestMethod]
    public void CopyRowFileTooManyParamsWrong()
    {
        Result result;
        LexerowCore core = new LexerowCore();

        // create a basic script
        List<string> lines = [
            "file= SelectFiles(\"data.xlsx\")",
            "fileRes= CreateExcel(\"result.xlsx\")",
            "CopyHeader(\"data.xlsx\", \"result.xlsx\")",
            "OnExcel file",
            "  ForEach Row",
            "    If A.Cell >10 Then",
            "      CopyRow(fileRes,\"hello\")",
            "    End if",
            "  Next",
            "End OnExcel"
            ];

        // load the script and compile it
        result = core.LoadLinesScript("script", lines);
        Assert.IsFalse(result.Res);
        Assert.AreEqual(ErrorCode.ParserFctParamCountWrong, result.ListError[0].ErrorCode);
    }

}
