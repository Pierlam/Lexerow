using FakeItEasy;
using Lexerow.Core.ScriptCompile.Parse;
using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;
using Lexerow.Core.System.ScriptDef;
using Lexerow.Core.Tests._05_Common;

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
        int numLine = 0;
        List<ScriptLineTokens> script = new List<ScriptLineTokens>();

        // build one line of tokens
        ScriptLineTokensTest.CreateOnExcelFileString(numLine++, script, "\"data.xlsx\"");

        // ForEach Row
        TestTokensBuilder.AddLineForEachRow(numLine++, script);

        // If A.Cell >10 Then A.Cell=10
        TestTokensBuilder.BuidIfColCellEqualIntThenSetColCellInt(3, script);

        // Next
        TestTokensBuilder.AddLineNext(numLine++, script);

        // End OnExcel not present!


        //==>just to check the content of the script
        var scriptCheck = TestTokens2ScriptBuilder.BuildScript(script);

        //==> Parse the script tokens
        var logger = A.Fake<IActivityLogger>();
        Parser sa = new Parser(logger);

        // parse the tokens
        ExecResult execResult = new ExecResult();
        bool res = sa.Process(execResult, script, out List<InstrBase> listInstr);

        //==> Check result
        Assert.IsFalse(res);
        Assert.AreEqual(ErrorCode.ParserTokenExpected, execResult.ListError[0].ErrorCode);
        Assert.AreEqual("OnExcel", execResult.ListError[0].Param);
    }
}