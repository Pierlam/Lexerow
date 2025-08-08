using Lexerow.Core.Scripts;
using Lexerow.Core.System;
using Lexerow.Core.System.Compilator;
using Lexerow.Core.Tests._05_Common;
using NPOI.SS.Formula.Functions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Tests.ScriptLexicalAnalyze;


/// <summary>
/// Test script lexical analyzer.
/// </summary>
[TestClass]
public class ScriptLexicalAnalyzerBasicTests
{
    /// <summary>
    /// file=OpenExcel("data.xslx")
    /// file; = ; OpenExcel; (; "data.xlsx"; )
    /// </summary>
    [TestMethod]
    public void TestBasic()
    {
        Script script = ScriptBuilder.Build("#comment", "file=OpenExcel(\"data.xslx\")");

        ExecResult execResult = new ExecResult();

        // analyse the source code, line by line
        LexicalAnalyzer.Process(execResult , script, out List<ScriptLineTokens> lt, new LexicalAnalyzerConfig());

        Assert.IsTrue(execResult.Result);

        // only 1 line, remove the line containing the comment
        Assert.AreEqual(1, lt.Count);

        // the second line has 6 source code tokens
        Assert.AreEqual(6, lt[0].ListScriptToken.Count);
        Assert.AreEqual("file", lt[0].ListScriptToken[0].Value);
        Assert.AreEqual("=", lt[0].ListScriptToken[1].Value);
        Assert.AreEqual("OpenExcel", lt[0].ListScriptToken[2].Value);
        Assert.AreEqual("(", lt[0].ListScriptToken[3].Value);
        Assert.AreEqual("\"data.xslx\"", lt[0].ListScriptToken[4].Value);
        Assert.AreEqual(ScriptTokenType.String, lt[0].ListScriptToken[4].ScriptTokenType);
        Assert.AreEqual(")", lt[0].ListScriptToken[5].Value);
    }

    [TestMethod]
    public void SourceScriptIsEmpty()
    {
        Script script = new Script("fileName");
        ExecResult execResult = new ExecResult();

        // analyse the source code, line by line
        LexicalAnalyzer.Process(execResult, script, out List<ScriptLineTokens> lt, new LexicalAnalyzerConfig());
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

        Script script = ScriptBuilder.Build("file=OpenExcel(\"data.xslx)");

        // analyse the source code, line by line
        LexicalAnalyzer.Process(execResult, script, out List<ScriptLineTokens> lt, new LexicalAnalyzerConfig());
        Assert.IsFalse(execResult.Result);

        Assert.AreEqual(ErrorCode.LexAnalyzeFoundSgtringBadFormatted, execResult.ListError[0].ErrorCode);

    }

    /// <summary>
    /// process a script having only comment -> nothing to compile!
    /// </summary>
    [TestMethod]
    public void SourceScriptHasOnlyComment()
    {
        Script script = ScriptBuilder.Build("#comment");
        ExecResult execResult = new ExecResult();

        // analyse the source code, line by line
        LexicalAnalyzer.Process(execResult, script, out List<ScriptLineTokens> lt, new LexicalAnalyzerConfig());
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
        Script script = ScriptBuilder.Build("OnExcel \"file.xlsx\"");
        ExecResult execResult = new ExecResult();

        // analyse the source code, line by line
        LexicalAnalyzer.Process(execResult, script, out List<ScriptLineTokens> lt, new LexicalAnalyzerConfig());
        Assert.IsTrue(execResult.Result);

        Assert.AreEqual(1, lt.Count);

        Assert.AreEqual(2, lt[0].ListScriptToken.Count);
        Assert.AreEqual("OnExcel", lt[0].ListScriptToken[0].Value);
        Assert.AreEqual("\"file.xlsx\"", lt[0].ListScriptToken[1].Value);
    }

    /// <summary>
    /// analyse a source code
    ///   ForEach Row
    /// The result: split tokens
    /// </summary>
    [TestMethod]
    public void ParseForEacRowOk()
    {
        Script script = ScriptBuilder.Build("  ForEach Row");

        ExecResult execResult = new ExecResult();

        // analyse the source code, line by line
        LexicalAnalyzer.Process(execResult, script, out List<ScriptLineTokens> lt, new LexicalAnalyzerConfig());
        Assert.IsTrue(execResult.Result);

        // only 1 line, remove the line containing the comment
        Assert.AreEqual(1, lt.Count);

        Assert.AreEqual(2, lt[0].ListScriptToken.Count);
        Assert.AreEqual("ForEach", lt[0].ListScriptToken[0].Value);
        Assert.AreEqual("Row", lt[0].ListScriptToken[1].Value);
    }

    /// <summary>
    /// analyse a source code
    ///   ForEach Row    #comment
    /// The result: split tokens
    /// </summary>
    [TestMethod]
    public void ParseForEacRowCommentOk()
    {
        Script script = ScriptBuilder.Build("  ForEach Row #comment");

        ExecResult execResult = new ExecResult();

        // analyse the source code, line by line
        LexicalAnalyzer.Process(execResult, script, out List<ScriptLineTokens> lt, new LexicalAnalyzerConfig());
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
        Script script = ScriptBuilder.Build("  If A.Cell>10 Then A.Cell=10");
        ExecResult execResult = new ExecResult();

        // analyse the source code, line by line
        LexicalAnalyzer.Process(execResult, script, out List<ScriptLineTokens> lt, new LexicalAnalyzerConfig());
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
}
