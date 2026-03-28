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
public class LoadFileExecCopyRowTests : BaseTests
{
    /// <summary>
    /// file=SelectFiles(".\10-ExcelFiles\copyRow1.xlsx")
    /// fileRes=CreateExcel(".\10-ExcelFiles\copyRowRes1.xlsx")
    /// CopyHeader(file, fileRes)
    /// OnExcel file
	///   ForEachRow
    ///     If A.Cell>10 Then CopyRow(fileRes)
	///   Next
    /// End OnExcel
    /// 
    /// Will copy the header and 2 datarow to the excel file result.
    /// </summary>     
    [TestMethod]
    public void CopyRowOk()
    {

        // remove the file if it exists
        if (File.Exists(PathExcelFilesExec + "copyRowRes1.xlsx"))
            File.Delete(PathExcelFilesExec + "copyRowRes1.xlsx");

        Result result;
        LexerowCore core = new LexerowCore();
        string scriptfile = PathScriptFiles + "copyRow1.lxrw";

        // load the script, compile it and then execute it
        result = core.LoadExecScript("script", scriptfile);
        Assert.IsTrue(result.Res);


        //--check the file is created
        ExcelFile excelFile = TestExcelChecker.Open(PathExcelFilesExec + "copyRowRes1.xlsx");
        Assert.IsNotNull(excelFile);

        // check the file has one sheet named xxx
        ExcelSheet excelSheet = TestExcelChecker.GetSheetByName(excelFile, CoreDefinitions.DefaultExcelSheetName);

        // check cells , row1
        bool res = TestExcelChecker.CheckCellValue(excelFile, "A1", "Id");
        Assert.IsTrue(res);
        res = TestExcelChecker.CheckCellValue(excelFile, "B1", "Name");
        Assert.IsTrue(res);
        res = TestExcelChecker.CheckCellValue(excelFile, "C1", "Cost");
        Assert.IsTrue(res);

        // the excel result should contains 2 data rows
        res = TestExcelChecker.CheckCellValue(excelFile, "A2", 2);
        Assert.IsTrue(res);
        res = TestExcelChecker.CheckCellValue(excelFile, "B2", "Henri");
        Assert.IsTrue(res);
        res = TestExcelChecker.CheckCellValue(excelFile, "C2", 10);
        Assert.IsTrue(res);

        // the excel result should contains 2 data rows
        res = TestExcelChecker.CheckCellValue(excelFile, "A3", 3);
        Assert.IsTrue(res);
        res = TestExcelChecker.CheckCellValue(excelFile, "B3", "Luis");
        Assert.IsTrue(res);
        res = TestExcelChecker.CheckCellValue(excelFile, "C3", 12);
        Assert.IsTrue(res);
    }

}
