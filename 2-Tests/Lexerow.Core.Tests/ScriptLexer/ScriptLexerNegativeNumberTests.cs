using FakeItEasy;
using Lexerow.Core.ScriptCompile.lex;
using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;
using Lexerow.Core.System.ScriptCompile;
using Lexerow.Core.System.ScriptDef;
using Lexerow.Core.Tests._05_Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Tests.ScriptLexer;

/// <summary>
/// Test script Lexer basic tests: lexical analyzer on SelectFiles instr.
/// </summary>
[TestClass]
public class ScriptLexerNegativeNumberTests
{
    /// <summary>
    /// </summary>
    [TestMethod]
    public void SetVaraSpaceMinus1Ok()
    {
        Script script = TestTokensBuilder.CreateScript("a= -1");

        //=>exec lexer, line by line
        Result result = new Result();
        Lexer.Process(A.Fake<IActivityLogger>(), result, script, out List<ScriptLineTokens> lt, new LexerConfig());

        //=> Check the result
        Assert.IsTrue(result.Res);
        Assert.AreEqual(1, lt.Count);
        Assert.AreEqual(4, lt[0].ListScriptToken.Count);
    }

    /// <summary>
    /// </summary>
    [TestMethod]
    public void SetVaraMinus1Ok()
    {
        Script script = TestTokensBuilder.CreateScript("a=-1");

        //=>exec lexer, line by line
        Result result = new Result();
        Lexer.Process(A.Fake<IActivityLogger>(), result, script, out List<ScriptLineTokens> lt, new LexerConfig());

        //=> Check the result
        Assert.IsTrue(result.Res);
        Assert.AreEqual(1, lt.Count);
        Assert.AreEqual(4, lt[0].ListScriptToken.Count);
    }

    /// <summary>
    /// </summary>
    [TestMethod]
    public void IfaEquMinus1ThenOk()
    {
        Script script = TestTokensBuilder.CreateScript("If a=-1 Then");

        //=>exec lexer, line by line
        Result result = new Result();
        Lexer.Process(A.Fake<IActivityLogger>(), result, script, out List<ScriptLineTokens> lt, new LexerConfig());

        //=> Check the result
        Assert.IsTrue(result.Res);
        Assert.AreEqual(1, lt.Count);
        Assert.AreEqual(6, lt[0].ListScriptToken.Count);
    }
}
