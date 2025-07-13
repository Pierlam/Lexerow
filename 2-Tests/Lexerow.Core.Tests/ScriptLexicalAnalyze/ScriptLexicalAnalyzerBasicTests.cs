using Lexerow.Core.Scripts;
using Lexerow.Core.System.Compilator;
using Lexerow.Core.Tests._05_Common;
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
    /// file= OpenExcel("data.xslx")
    /// file; = ; OpenExcel; (; "data.xlsx"; )
    /// </summary>
    [TestMethod]
    public void TestBasic()
    {
        SourceScript sourceScript = SourceScriptBuilder.Build("fileName");

        // analyse the source code, line by line
        LexicalAnalyzer.Process(sourceScript, out List<SourceCodeLineTokens> listSourceCodeLineTokens, new LexicalAnalyzerConfig());

        // only 1 line, remove the line containing the comment
        Assert.AreEqual(1, listSourceCodeLineTokens.Count);

        // the second line has 6 source code tokens
        Assert.AreEqual(6, listSourceCodeLineTokens[0].ListSourceCodeToken.Count);
        Assert.AreEqual("file", listSourceCodeLineTokens[0].ListSourceCodeToken[0].Value);
        Assert.AreEqual("=", listSourceCodeLineTokens[0].ListSourceCodeToken[1].Value);
        Assert.AreEqual("OpenExcel", listSourceCodeLineTokens[0].ListSourceCodeToken[2].Value);
        Assert.AreEqual("(", listSourceCodeLineTokens[0].ListSourceCodeToken[3].Value);
        Assert.AreEqual("\"data.xslx\"", listSourceCodeLineTokens[0].ListSourceCodeToken[4].Value);
        Assert.AreEqual(")", listSourceCodeLineTokens[0].ListSourceCodeToken[5].Value);
    }

    [TestMethod]
    public void SourceScriptIsEmpty()
    {
        SourceScript sourceScript = new SourceScript("fileName");

        // analyse the source code, line by line
        LexicalAnalyzer.Process(sourceScript, out List<SourceCodeLineTokens> listSourceCodeLineTokens, new LexicalAnalyzerConfig());

        Assert.AreEqual(0, listSourceCodeLineTokens.Count);
    }

    /// <summary>
    /// process a script having only comment -> nothing to compile!
    /// </summary>
    [TestMethod]
    public void SourceScriptHasOnlyComment()
    {
        SourceScript sourceScript = SourceScriptBuilder.BuildOneLineComment("fileName");

        // analyse the source code, line by line
        LexicalAnalyzer.Process(sourceScript, out List<SourceCodeLineTokens> listSourceCodeLineTokens, new LexicalAnalyzerConfig());

        Assert.AreEqual(0, listSourceCodeLineTokens.Count);
    }

    // TODO: test specific source codes!
    //ici()
}
