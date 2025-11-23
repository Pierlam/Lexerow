using FakeItEasy;
using Lexerow.Core.ScriptCompile.lex;
using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;
using Lexerow.Core.System.ScriptCompile;
using Lexerow.Core.System.ScriptDef;
using Lexerow.Core.Tests._05_Common;

namespace Lexerow.Core.Tests.ScriptLexer;

/// <summary>
/// Test script lexical analyzer.
/// </summary>
[TestClass]
public class ScriptLexerBasicTests
{
    /// <summary>
    /// file=SelectFiles("data.xslx")
    /// file; = ; SelectFiles; (; "data.xlsx"; )
    /// </summary>
    [TestMethod]
    public void SelectFilesOk()
    {
        Script script = TestTokensBuilder.CreateScript("#comment", "file=SelectFiles(\"data.xslx\")");

        //=>exec lexer, line by line
        Result result = new Result();
        Lexer.Process(A.Fake<IActivityLogger>(), result, script, out List<ScriptLineTokens> lt, new LexerConfig());

        //=> Check the result
        Assert.IsTrue(result.Res);

        // only 1 line, remove the line containing the comment
        Assert.AreEqual(1, lt.Count);

        // the line has 6 source code tokens
        Assert.AreEqual(6, lt[0].ListScriptToken.Count);
        Assert.AreEqual("file", lt[0].ListScriptToken[0].Value);
        Assert.AreEqual("=", lt[0].ListScriptToken[1].Value);
        Assert.AreEqual("SelectFiles", lt[0].ListScriptToken[2].Value);
        Assert.AreEqual("(", lt[0].ListScriptToken[3].Value);
        Assert.AreEqual("\"data.xslx\"", lt[0].ListScriptToken[4].Value);
        Assert.AreEqual(ScriptTokenType.String, lt[0].ListScriptToken[4].ScriptTokenType);
        Assert.AreEqual(")", lt[0].ListScriptToken[5].Value);
    }

    [TestMethod]
    public void SourceScriptIsEmpty()
    {
        Script script = new Script("name", "fileName");

        //=>exec lexer, line by line
        Result result = new Result();
        Lexer.Process(A.Fake<IActivityLogger>(), result, script, out List<ScriptLineTokens> lt, new LexerConfig());

        //=> Check the result
        Assert.IsTrue(result.Res);

        Assert.AreEqual(0, lt.Count);
    }

    /// <summary>
    /// the source code line is bad formatted, the string end tag is missing.
    /// file=OpenExcel("data.xslx)
    /// </summary>
    [TestMethod]
    public void SourceTokenStringEndTagMissingErr()
    {
        Script script = TestTokensBuilder.CreateScript("file=SelectFiles(\"data.xslx)");

        //=>exec lexer, line by line
        Result result = new Result();
        Lexer.Process(A.Fake<IActivityLogger>(), result, script, out List<ScriptLineTokens> lt, new LexerConfig());

        //=> Check the result
        Assert.IsFalse(result.Res);

        Assert.AreEqual(ErrorCode.LexerFoundSgtringBadFormatted, result.ListError[0].ErrorCode);
    }

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

    /// <summary>
    /// analyse a source code
    /// OnExcel "file.xlsx"
    ///
    /// The result: split tokens
    /// </summary>
    [TestMethod]
    public void ParseOnExcelFilenameOk()
    {
        Script script = TestTokensBuilder.CreateScript("OnExcel \"file.xlsx\"");

        //=>exec lexer, line by line
        Result result = new Result();
        Lexer.Process(A.Fake<IActivityLogger>(), result, script, out List<ScriptLineTokens> lt, new LexerConfig());

        //=> Check the result
        Assert.IsTrue(result.Res);

        Assert.AreEqual(1, lt.Count);

        Assert.AreEqual(2, lt[0].ListScriptToken.Count);
        Assert.AreEqual("OnExcel", lt[0].ListScriptToken[0].Value);
        Assert.AreEqual(ScriptTokenType.Name, lt[0].ListScriptToken[0].ScriptTokenType);

        Assert.AreEqual("\"file.xlsx\"", lt[0].ListScriptToken[1].Value);
        Assert.AreEqual(ScriptTokenType.String, lt[0].ListScriptToken[1].ScriptTokenType);
    }

    /// <summary>
    /// analyse a source code
    ///   ForEach Row
    /// The result: split tokens
    /// </summary>
    [TestMethod]
    public void ParseForEachRowOk()
    {
        Script script = TestTokensBuilder.CreateScript("  ForEach Row");

        //=>exec lexer, line by line
        Result result = new Result();
        Lexer.Process(A.Fake<IActivityLogger>(), result, script, out List<ScriptLineTokens> lt, new LexerConfig());

        //=> Check the result
        Assert.IsTrue(result.Res);

        // only 1 line, remove the line containing the comment
        Assert.AreEqual(1, lt.Count);

        Assert.AreEqual(2, lt[0].ListScriptToken.Count);
        Assert.AreEqual("ForEach", lt[0].ListScriptToken[0].Value);
        Assert.AreEqual(ScriptTokenType.Name, lt[0].ListScriptToken[0].ScriptTokenType);
        Assert.AreEqual("Row", lt[0].ListScriptToken[1].Value);
        Assert.AreEqual(ScriptTokenType.Name, lt[0].ListScriptToken[1].ScriptTokenType);
    }

    /// <summary>
    /// analyse a source code
    ///   ForEach Row    #comment
    /// The result: split tokens
    /// </summary>
    [TestMethod]
    public void ParseForEachRowCommentOk()
    {
        Script script = TestTokensBuilder.CreateScript("  ForEach Row #comment");

        //=>exec lexer, line by line
        Result result = new Result();
        Lexer.Process(A.Fake<IActivityLogger>(), result, script, out List<ScriptLineTokens> lt, new LexerConfig());

        //=> Check the result
        Assert.IsTrue(result.Res);

        // only 1 line, remove the line containing the comment
        Assert.AreEqual(1, lt.Count);

        Assert.AreEqual(2, lt[0].ListScriptToken.Count);
        Assert.AreEqual("ForEach", lt[0].ListScriptToken[0].Value);
        Assert.AreEqual("Row", lt[0].ListScriptToken[1].Value);
    }

    /// <summary>
    /// analyse the source code:
    ///      If A.Cell >10 Then A.Cell= 10
    ///
    /// The result: split tokens
    /// </summary>
    [TestMethod]
    public void ParseIfACellGt10ThenACellEq10Ok()
    {
        Script script = TestTokensBuilder.CreateScript("If A.Cell>10 Then A.Cell=10");

        //=>exec lexer, line by line
        Result result = new Result();
        Lexer.Process(A.Fake<IActivityLogger>(), result, script, out List<ScriptLineTokens> lt, new LexerConfig());

        //=> Check the result
        Assert.IsTrue(result.Res);

        Assert.AreEqual(1, lt.Count);

        Assert.AreEqual(12, lt[0].ListScriptToken.Count);
        Assert.AreEqual("If", lt[0].ListScriptToken[0].Value);
        Assert.AreEqual("A", lt[0].ListScriptToken[1].Value);

        Assert.AreEqual(".", lt[0].ListScriptToken[2].Value);
        Assert.AreEqual(ScriptTokenType.Separator, lt[0].ListScriptToken[2].ScriptTokenType);

        Assert.AreEqual("Cell", lt[0].ListScriptToken[3].Value);
        Assert.AreEqual(">", lt[0].ListScriptToken[4].Value);
        Assert.AreEqual(ScriptTokenType.Separator, lt[0].ListScriptToken[4].ScriptTokenType);

        Assert.AreEqual("10", lt[0].ListScriptToken[5].Value);
        Assert.AreEqual(ScriptTokenType.Integer, lt[0].ListScriptToken[5].ScriptTokenType);
        Assert.AreEqual("Then", lt[0].ListScriptToken[6].Value);
        Assert.AreEqual("A", lt[0].ListScriptToken[7].Value);
        Assert.AreEqual(".", lt[0].ListScriptToken[8].Value);
        Assert.AreEqual("Cell", lt[0].ListScriptToken[9].Value);
        Assert.AreEqual("=", lt[0].ListScriptToken[10].Value);
        Assert.AreEqual(ScriptTokenType.Separator, lt[0].ListScriptToken[10].ScriptTokenType);

        Assert.AreEqual("10", lt[0].ListScriptToken[11].Value);
        Assert.AreEqual(ScriptTokenType.Integer, lt[0].ListScriptToken[11].ScriptTokenType);
    }

    [TestMethod]
    public void ParseIfACellEqDateOk()
    {
        Script script = TestTokensBuilder.CreateScript("If A.Cell=Date(2025,11,23)");

        //=>exec lexer, line by line
        Result result = new Result();
        Lexer.Process(A.Fake<IActivityLogger>(), result, script, out List<ScriptLineTokens> listlineTokens, new LexerConfig());

        //=> Check the result
        Assert.IsTrue(result.Res);
        Assert.AreEqual(1, listlineTokens.Count);

        var lineTokens = listlineTokens[0];
        Assert.AreEqual(13, lineTokens.ListScriptToken.Count);

        int idx = 5;
        Assert.IsTrue(TestTokensHelper.TestName(lineTokens, idx++, "Date"));
        Assert.IsTrue(TestTokensHelper.TestSep(lineTokens, idx++, "("));
        Assert.IsTrue(TestTokensHelper.TestNumberInt(lineTokens, idx++, 2025));
        Assert.IsTrue(TestTokensHelper.TestSep(lineTokens, idx++, ","));
        Assert.IsTrue(TestTokensHelper.TestNumberInt(lineTokens, idx++, 11));
        Assert.IsTrue(TestTokensHelper.TestSep(lineTokens, idx++, ","));
        Assert.IsTrue(TestTokensHelper.TestNumberInt(lineTokens, idx++, 23));
        Assert.IsTrue(TestTokensHelper.TestSep(lineTokens, idx++, ")"));

    }
}