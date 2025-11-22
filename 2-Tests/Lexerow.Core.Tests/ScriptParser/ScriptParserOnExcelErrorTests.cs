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
    ///	OnExcel "data.xlsx"
    ///   ForEach Row
    ///     If A.Cell >10 Then A.Cell=10
    ///   Next
    ///         -> End OnExcel is missing ! -> error
    /// </summary>
    [TestMethod]
    public void EndOnExcelMissingErr()
    {
        int numLine = 0;
        List<ScriptLineTokens> scriptTokens = new List<ScriptLineTokens>();

        // OnExcel "data.xlsx"
        TestTokensBuilder.AddLineOnExcelFileString(numLine++, scriptTokens, "\"data.xlsx\"");

        // ForEach Row
        TestTokensBuilder.AddLineForEachRow(numLine++, scriptTokens);

        // If A.Cell >10 Then A.Cell=10
        TestTokensBuilder.BuidIfColCellEqualIntThenSetColCellInt(3, scriptTokens);

        // Next
        TestTokensBuilder.AddLineNext(numLine++, scriptTokens);

        // End OnExcel not present!


        //==>just to check the content of the script
        var scriptCheck = TestTokens2ScriptBuilder.BuildScript(scriptTokens);

        //==> Parse the script tokens
        Parser parser = new Parser(A.Fake<IActivityLogger>());
        ExecResult execResult = new ExecResult();
        var prog = TestInstrBuilder.CreateProgram();
        bool res = parser.Process(execResult, scriptTokens, prog);

        //==> Check the result
        Assert.IsFalse(res);
        Assert.AreEqual(ErrorCode.ParserTokenExpected, execResult.ListError[0].ErrorCode);
        Assert.AreEqual("OnExcel", execResult.ListError[0].Param);
    }
}