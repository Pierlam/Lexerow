using FakeItEasy;
using Lexerow.Core.ScriptCompile.Parse;
using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;
using Lexerow.Core.System.ScriptDef;
using Lexerow.Core.Tests._05_Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Tests.ScriptParser;

/// <summary>
/// Test script parser on OnExcel instr.
/// Focus on FirstRow instr.
/// </summary>
[TestClass]
public class ScriptParserOnExcelFirstRowOkTests
{
    /// <summary>
    /// Default: sheet=0
    ///
    ///	OnExcel "file.xlsx"
    ///	  FirstRow 3
    ///   ForEach Row
    ///     If A.Cell>10 Then A.Cell=10
    ///   Next
    /// End OnExcel
    /// </summary>
    [TestMethod]
    public void OnExcelFirstRowOk()
    {
        int numLine = 1;
        List<ScriptLineTokens> script = new List<ScriptLineTokens>();

        //-build one line of tokens
        ScriptLineTokensTest.CreateOnExcelFileString(numLine++ , script, "\"data.xlsx\"");

        // FirstRow 3
        TestTokensBuilder.AddLineFirstRow(numLine++, script, 3);

        // ForEach Row
        TestTokensBuilder.AddLineForEachRow(numLine++, script);

        // If A.Cell>10 Then A.Cell=10
        TestTokensBuilder.BuidIfColCellCompIntThenSetColCellInt(numLine++, script, "A", ">", 10, "A", 10);

        // Next
        TestTokensBuilder.AddLineNext(numLine++, script);

        // End OnExcel
        TestTokensBuilder.AddLineEndOnExcel(numLine++, script);

        //==>just to check the content of the script
        var scriptCheck = TestTokens2ScriptBuilder.BuildScript(script);

        //==> Parse the script tokens
        var logger = A.Fake<IActivityLogger>();
        Parser sa = new Parser(logger);
        ExecResult execResult = new ExecResult();
        bool res = sa.Process(execResult, script, out List<InstrBase> listInstr);

        //==> check the result
        Assert.IsTrue(res);
        Assert.AreEqual(1, listInstr.Count);

        // OnExcel
        Assert.AreEqual(InstrType.OnExcel, listInstr[0].InstrType);
        InstrOnExcel instrOnExcel = listInstr[0] as InstrOnExcel;

        // check the firstrow value
        Assert.AreEqual(3, instrOnExcel.ListSheets[0].FirstRowNum);  
    }

    /// <summary>
    /// Default: sheet=0
    ///
    ///	OnExcel "file.xlsx"
    ///	  FirstRow 0
    ///   ForEach Row
    ///     If A.Cell>10 Then A.Cell=10
    ///   Next
    /// End OnExcel
    /// </summary>
    [TestMethod]
    public void OnExcelFirstRowZeroError()
    {
        // FirstRow 0
        //ici();
    }

    // FirstRow a

}
