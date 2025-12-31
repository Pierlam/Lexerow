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
/// Test script lexical analyzer.
/// </summary>
[TestClass]
public class ScriptLexerBasicTests
{
    /// <summary>
    /// process a script having only comment -> nothing to compile!
    /// </summary>
    [TestMethod]
    public void SourceScriptHasOnlyComment()
    {
        Script script = TestTokensBuilder.CreateScript("#comment");

        //=>exec lexer, line by line
        Result result = new Result();
        Lexer.Process(A.Fake<IActivityLogger>(), result, script, out List<ScriptLineTokens> lt, new LexerConfig());

        //=> Check the result
        Assert.IsTrue(result.Res);

        Assert.AreEqual(0, lt.Count);
    }

    [TestMethod]
    public void ParseVarOk()
    {
        Script script = TestTokensBuilder.CreateScript("myvar");

        //=>exec lexer, line by line
        Result result = new Result();
        Lexer.Process(A.Fake<IActivityLogger>(), result, script, out List<ScriptLineTokens> lt, new LexerConfig());

        //=> Check the result
        Assert.IsTrue(result.Res);
        Assert.AreEqual(1, lt.Count);

        var lineTokens = lt[0];
        Assert.AreEqual(1, lineTokens.ListScriptToken.Count);

        int idx = 0;
        Assert.IsTrue(TestTokensHelper.TestName(lineTokens, idx++, "myvar"));
    }

    [TestMethod]
    public void ParseVar12Ok()
    {
        Script script = TestTokensBuilder.CreateScript("my12_var");

        //=>exec lexer, line by line
        Result result = new Result();
        Lexer.Process(A.Fake<IActivityLogger>(), result, script, out List<ScriptLineTokens> lt, new LexerConfig());

        //=> Check the result
        Assert.IsTrue(result.Res);
        Assert.AreEqual(1, lt.Count);

        var lineTokens = lt[0];
        Assert.AreEqual(1, lineTokens.ListScriptToken.Count);

        int idx = 0;
        Assert.IsTrue(TestTokensHelper.TestName(lineTokens, idx++, "my12_var"));
    }

    [TestMethod]
    public void Parse12VarWrong()
    {
        Script script = TestTokensBuilder.CreateScript("12myvar");

        //=>exec lexer, line by line
        Result result = new Result();
        Lexer.Process(A.Fake<IActivityLogger>(), result, script, out List<ScriptLineTokens> listlineTokens, new LexerConfig());

        //=> Check the result
        Assert.IsFalse(result.Res);

        Assert.AreEqual(ErrorCode.LexerFoundDoubleWrong, result.ListError[0].ErrorCode);
    }


    [TestMethod]
    public void ParseVarSystWrong()
    {
        Script script = TestTokensBuilder.CreateScript("$ DateFormat");

        //=>exec lexer, line by line
        Result result = new Result();
        Lexer.Process(A.Fake<IActivityLogger>(), result, script, out List<ScriptLineTokens> listlineTokens, new LexerConfig());

        //=> Check the result
        Assert.IsFalse(result.Res);

        Assert.AreEqual(ErrorCode.LexerFoundSystNameWrong, result.ListError[0].ErrorCode);
    }

    [TestMethod]
    public void ParseVarSystAloneWrong()
    {
        Script script = TestTokensBuilder.CreateScript("$");

        //=>exec lexer, line by line
        Result result = new Result();
        Lexer.Process(A.Fake<IActivityLogger>(), result, script, out List<ScriptLineTokens> listlineTokens, new LexerConfig());

        //=> Check the result
        Assert.IsFalse(result.Res);

        Assert.AreEqual(ErrorCode.LexerFoundSystNameWrong, result.ListError[0].ErrorCode);
    }

    [TestMethod]
    public void ParseVarSystAtEndWrong()
    {
        Script script = TestTokensBuilder.CreateScript("myvar $");

        //=>exec lexer, line by line
        Result result = new Result();
        Lexer.Process(A.Fake<IActivityLogger>(), result, script, out List<ScriptLineTokens> listlineTokens, new LexerConfig());

        //=> Check the result
        Assert.IsFalse(result.Res);

        Assert.AreEqual(ErrorCode.LexerFoundSystNameWrong, result.ListError[0].ErrorCode);
    }

    [TestMethod]
    public void ParseVarSystStringWrong()
    {
        Script script = TestTokensBuilder.CreateScript("$\"str\"");

        //=>exec lexer, line by line
        Result result = new Result();
        Lexer.Process(A.Fake<IActivityLogger>(), result, script, out List<ScriptLineTokens> listlineTokens, new LexerConfig());

        //=> Check the result
        Assert.IsFalse(result.Res);

        Assert.AreEqual(ErrorCode.LexerFoundSystNameWrong, result.ListError[0].ErrorCode);
    }

    [TestMethod]
    public void ParseVarSystIntegerWrong()
    {
        Script script = TestTokensBuilder.CreateScript("$12");

        //=>exec lexer, line by line
        Result result = new Result();
        Lexer.Process(A.Fake<IActivityLogger>(), result, script, out List<ScriptLineTokens> listlineTokens, new LexerConfig());

        //=> Check the result
        Assert.IsFalse(result.Res);

        Assert.AreEqual(ErrorCode.LexerFoundSystNameWrong, result.ListError[0].ErrorCode);
    }

}
