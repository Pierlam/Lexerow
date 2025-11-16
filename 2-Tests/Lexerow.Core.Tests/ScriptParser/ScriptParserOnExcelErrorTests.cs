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
/// Test script lexical analyzer on OnExcel instr having error.
/// Neither OnSheet, nor FirstRow.
/// One IfThen in ForEachRow.
/// </summary>
[TestClass]
public class ScriptParserOnExcelErrorTests
{
    /// <summary>
    /// Compile: OnExcel, very short version
    /// Result: one instruction OnExcel
    /// Implicite: sheet=0, FirstRow=1
    /// 
    ///	OnExcel "file.xlsx"
    ///   ForEach Row
    ///     If A.Cell >10 Then A.Cell=10
    ///   Next 
    ///         -> End OnExcel is missing ! -> error
    /// </summary>
    [TestMethod]
    public void EndOnExcelMissingErr()
    {
        ScriptLineTokensTest lineTok;
        List<ScriptLineTokens> script = new List<ScriptLineTokens>();

        //-build one line of tokens
        lineTok = ScriptLineTokensTest.CreateOnExcelFileString("\"data.xlsx\"");
        script.Add(lineTok);

        // ForEach Row
        lineTok = new ScriptLineTokensTest();
        lineTok.AddTokenName(2, "ForEach", "Row");
        script.Add(lineTok);

        // If A.Cell >10 Then A.Cell=10
        TestTokensBuilder.BuidIfColCellEqualIntThenSetColCellInt(3, script);

        // Next
        lineTok = new ScriptLineTokensTest();
        lineTok.AddTokenName(1, 1, "Next");
        script.Add(lineTok);

        // End OnExcel
        //lineTok = new ScriptLineTokensTest();
        //lineTok.AddTokenName(1, "End", "OnExcel");
        //script.Add(lineTok);

        var logger = A.Fake<IActivityLogger>();
        Parser sa = new Parser(logger);

        //--parse the tokens
        ExecResult execResult = new ExecResult();
        bool res = sa.Process(execResult, script, out List<InstrBase> listInstr);

        Assert.IsFalse(res);
        Assert.AreEqual(ErrorCode.ParserTokenExpected, execResult.ListError[0].ErrorCode);
        Assert.AreEqual("OnExcel", execResult.ListError[0].Param);

    }
}
