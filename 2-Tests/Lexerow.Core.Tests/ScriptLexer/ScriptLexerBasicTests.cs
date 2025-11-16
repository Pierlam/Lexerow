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
        Script script = TestTokensBuilder.Build("#comment", "file=SelectFiles(\"data.xslx\")");

        ExecResult execResult = new ExecResult();

        // analyse the source code, line by line
        var logger = A.Fake<IActivityLogger>();
        Lexer.Process(logger, execResult , script, out List<ScriptLineTokens> lt, new LexerConfig());

        Assert.IsTrue(execResult.Result);

        // only 1 line, remove the line containing the comment
        Assert.AreEqual(1, lt.Count);

        // the second line has 6 source code tokens
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
        ExecResult execResult = new ExecResult();

        var logger = A.Fake<IActivityLogger>();

        // analyse the source code, line by line
        Lexer.Process(logger, execResult, script, out List<ScriptLineTokens> lt, new LexerConfig());
        Assert.IsTrue(execResult.Result);

        Assert.AreEqual(0, lt.Count);
    }

    /// <summary>
    /// the source code line is bad formatted, the string end tag is missing.
    /// file=OpenExcel("data.xslx)
    /// </summary>
    [TestMethod]
    public void SourceTokenStringEndTagMissingErr()
    {
        ExecResult execResult = new ExecResult();

        Script script = TestTokensBuilder.Build("file=SelectFiles(\"data.xslx)");

        var logger = A.Fake<IActivityLogger>();

        // analyse the source code, line by line
        Lexer.Process(logger, execResult, script, out List<ScriptLineTokens> lt, new LexerConfig());
        Assert.IsFalse(execResult.Result);

        Assert.AreEqual(ErrorCode.LexerFoundSgtringBadFormatted, execResult.ListError[0].ErrorCode);

    }

    /// <summary>
    /// process a script having only comment -> nothing to compile!
    /// </summary>
    [TestMethod]
    public void SourceScriptHasOnlyComment()
    {
        Script script = TestTokensBuilder.Build("#comment");
        ExecResult execResult = new ExecResult();

        var logger = A.Fake<IActivityLogger>();

        // analyse the source code, line by line
        Lexer.Process(logger, execResult, script, out List<ScriptLineTokens> lt, new LexerConfig());
        Assert.IsTrue(execResult.Result);

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
        Script script = TestTokensBuilder.Build("OnExcel \"file.xlsx\"");
        ExecResult execResult = new ExecResult();

        var logger = A.Fake<IActivityLogger>();

        // analyse the source code, line by line
        Lexer.Process(logger, execResult, script, out List<ScriptLineTokens> lt, new LexerConfig());
        Assert.IsTrue(execResult.Result);

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
        Script script = TestTokensBuilder.Build("  ForEach Row");

        ExecResult execResult = new ExecResult();

        var logger = A.Fake<IActivityLogger>();

        // analyse the source code, line by line
        Lexer.Process(logger, execResult, script, out List<ScriptLineTokens> lt, new LexerConfig());
        Assert.IsTrue(execResult.Result);

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
        Script script = TestTokensBuilder.Build("  ForEach Row #comment");

        ExecResult execResult = new ExecResult();

        var logger = A.Fake<IActivityLogger>();

        // analyse the source code, line by line
        Lexer.Process(logger, execResult, script, out List<ScriptLineTokens> lt, new LexerConfig());
        Assert.IsTrue(execResult.Result);

        // only 1 line, remove the line containing the comment
        Assert.AreEqual(1, lt.Count);

        Assert.AreEqual(2, lt[0].ListScriptToken.Count);
        Assert.AreEqual("ForEach", lt[0].ListScriptToken[0].Value);
        Assert.AreEqual("Row", lt[0].ListScriptToken[1].Value);
    }


    /// <summary>
    /// analyse a source code
    ///      If A.Cell >10 Then A.Cell= 10
    /// 
    /// The result: split tokens
    /// </summary>
    [TestMethod]
    public void ParseIfACellGt10ThenACellEq10Ok()
    {
        Script script = TestTokensBuilder.Build("If A.Cell>10 Then A.Cell=10");
        ExecResult execResult = new ExecResult();

        var logger = A.Fake<IActivityLogger>();

        // analyse the source code, line by line
        Lexer.Process(logger, execResult, script, out List<ScriptLineTokens> lt, new LexerConfig());
        Assert.IsTrue(execResult.Result);

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

    // TODO: script is on several lines
}
