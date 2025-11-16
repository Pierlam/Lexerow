using Lexerow.Core.ExcelLayer;
using Lexerow.Core.System;
using Lexerow.Core.System.Excel;
using Lexerow.Core.Tests.Common;
using NPOI.SS.Formula.Functions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Tests.ExcelProcessor;

/// <summary>
/// IExcelProcessor, test Open and close excel file.
/// </summary>
[TestClass]
public class ExcelProcessorOpenCloseFileTests : BaseTests
{
    [TestMethod]
    public void BasicOpenExcelOk()
    {
        // injecter la lib NPOI
        IExcelProcessor proc = new ExcelProcessorNpoi();

        string fileName = PathExcelFilesExec + "OpenClose.xlsx";

        bool res = proc.Open(fileName, out IExcelFile excelFile, out ExecResultError coreError);
        Assert.IsTrue(res);
        Assert.IsNull(coreError);
        Assert.IsNotNull(excelFile);
        Assert.AreEqual(fileName, excelFile.FileName);

        res = proc.Close(excelFile, out ExecResultError error);
        Assert.IsTrue(res);
        Assert.IsNull(error);
    }

    [TestMethod]
    public void BasicOpenExcelDoestExistsFail()
    {
        // injecter la lib NPOI
        IExcelProcessor proc = new ExcelProcessorNpoi();

        string fileName = @"..\..\..\10-Files\Blabla.xlsx";

        bool res = proc.Open(fileName, out IExcelFile excelFile, out ExecResultError coreError);
        Assert.IsFalse(res);
        Assert.IsNull(excelFile);
        Assert.IsNotNull(coreError);
        Assert.AreEqual(ErrorCode.FileNotFound, coreError.ErrorCode);
    }
}
