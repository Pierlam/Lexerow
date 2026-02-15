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
    /// file= CreateExcel("mydata.xlsx")
    /// </summary>
    [TestMethod]
    public void CopyHeaderExcelFileOk()
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

    // CopyHeader(filename, excelOut)
    /// <summary>
    /// file= CreateExcel("mydata.xlsx")
    /// </summary>
    [TestMethod]
    public void CopyHeaderStringOk()
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
}
