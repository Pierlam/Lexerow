using Lexerow.Core.ScriptCompile.lex;
using Lexerow.Core.System.Compilator;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Tests.StringParser;

[TestClass]
public class StringParser
{
    [TestMethod]
    public void TestOneInt12Ok()
    {
        LexerConfig conf= new LexerConfig();

        ScriptSplitter stringParser = new ScriptSplitter();
        string line = "12";
        bool res = stringParser.Split(1, line, conf.Separators, conf.StringSep, conf.CommentTag, out List<ScriptToken> listScriptTokens, out ScriptTokenType lastTokenType);

        Assert.IsTrue(res);
        Assert.AreEqual(1, listScriptTokens.Count);
        Assert.AreEqual(ScriptTokenType.Integer, listScriptTokens[0].ScriptTokenType);
        Assert.AreEqual(12, listScriptTokens[0].ValueInt);
    }

    [TestMethod]
    public void TestOneInt12SpcOk()
    {
        LexerConfig conf = new LexerConfig();

        ScriptSplitter stringParser = new ScriptSplitter();
        string line = "12 ";
        bool res = stringParser.Split(1, line, conf.Separators, conf.StringSep, conf.CommentTag, out List<ScriptToken> listScriptTokens, out ScriptTokenType lastTokenType);

        Assert.IsTrue(res);
        Assert.AreEqual(1, listScriptTokens.Count);
        Assert.AreEqual(ScriptTokenType.Integer, listScriptTokens[0].ScriptTokenType);
        Assert.AreEqual(12, listScriptTokens[0].ValueInt);
    }

    [TestMethod]
    public void TestOneDouble_43dot95E10Ok()
    {
        LexerConfig conf = new LexerConfig();

        ScriptSplitter stringParser = new ScriptSplitter();
        string line = "43E10";
        bool res = stringParser.Split(1, line, conf.Separators, conf.StringSep, conf.CommentTag, out List<ScriptToken> listScriptTokens, out ScriptTokenType lastTokenType);

        Assert.IsTrue(res);
        Assert.AreEqual(1, listScriptTokens.Count);
        Assert.AreEqual(ScriptTokenType.Double, listScriptTokens[0].ScriptTokenType);
        Assert.AreEqual(43E10, listScriptTokens[0].ValueDouble);
    }


    [TestMethod]
    public void TestOneDouble_12dot45Ok()
    {
        LexerConfig conf = new LexerConfig();

        ScriptSplitter stringParser = new ScriptSplitter();
        string line = "12.45";
        bool res = stringParser.Split(1, line, conf.Separators, conf.StringSep, conf.CommentTag, out List<ScriptToken> listScriptTokens, out ScriptTokenType lastTokenType);

        Assert.IsTrue(res);
        Assert.AreEqual(1, listScriptTokens.Count);
        Assert.AreEqual(ScriptTokenType.Double, listScriptTokens[0].ScriptTokenType);
        Assert.AreEqual(12.45, listScriptTokens[0].ValueDouble);
    }

    [TestMethod]
    public void TestOneDouble_dot789Ok()
    {
        LexerConfig conf = new LexerConfig();

        ScriptSplitter stringParser = new ScriptSplitter();
        string line = ".789";
        bool res = stringParser.Split(1, line, conf.Separators, conf.StringSep, conf.CommentTag, out List<ScriptToken> listScriptTokens, out ScriptTokenType lastTokenType);

        Assert.IsTrue(res);
        Assert.AreEqual(1, listScriptTokens.Count);
        Assert.AreEqual(ScriptTokenType.Double, listScriptTokens[0].ScriptTokenType);
        Assert.AreEqual(.789, listScriptTokens[0].ValueDouble);
    }

    [TestMethod]
    public void TestOneDouble_4dot95E10Ok()
    {
        LexerConfig conf = new LexerConfig();

        ScriptSplitter stringParser = new ScriptSplitter();
        string line = "4.95E10";
        bool res = stringParser.Split(1, line, conf.Separators, conf.StringSep, conf.CommentTag, out List<ScriptToken> listScriptTokens, out ScriptTokenType lastTokenType);

        Assert.IsTrue(res);
        Assert.AreEqual(1, listScriptTokens.Count);
        Assert.AreEqual(ScriptTokenType.Double, listScriptTokens[0].ScriptTokenType);
        Assert.AreEqual(4.95E10, listScriptTokens[0].ValueDouble);
    }

    [TestMethod]
    public void TestOneDouble_4dot95Eminus10Ok()
    {
        LexerConfig conf = new LexerConfig();

        ScriptSplitter stringParser = new ScriptSplitter();
        string line = "4.95E-10";
        bool res = stringParser.Split(1, line, conf.Separators, conf.StringSep, conf.CommentTag, out List<ScriptToken> listScriptTokens, out ScriptTokenType lastTokenType);

        Assert.IsTrue(res);
        Assert.AreEqual(1, listScriptTokens.Count);
        Assert.AreEqual(ScriptTokenType.Double, listScriptTokens[0].ScriptTokenType);
        Assert.AreEqual(4.95E-10, listScriptTokens[0].ValueDouble);
    }

    [TestMethod]
    public void TestOneIntWrong()
    {
        LexerConfig conf = new LexerConfig();

        ScriptSplitter stringParser = new ScriptSplitter();
        string line = " 12A";
        bool res = stringParser.Split(1, line, conf.Separators, conf.StringSep, conf.CommentTag, out List<ScriptToken> listScriptTokens, out ScriptTokenType lastTokenType);

        // return false -> error
        Assert.IsFalse(res);
        Assert.AreEqual(1, listScriptTokens.Count);
        Assert.AreEqual(ScriptTokenType.WrongNumber, listScriptTokens[0].ScriptTokenType);
    }
    [TestMethod]
    public void TestOneIntWrong_spc()
    {
        LexerConfig conf = new LexerConfig();

        ScriptSplitter stringParser = new ScriptSplitter();
        string line = " 12A ";
        bool res = stringParser.Split(1, line, conf.Separators, conf.StringSep, conf.CommentTag, out List<ScriptToken> listScriptTokens, out ScriptTokenType lastTokenType);

        // return false -> error
        Assert.IsFalse(res);
        Assert.AreEqual(1, listScriptTokens.Count);
        Assert.AreEqual(ScriptTokenType.WrongNumber, listScriptTokens[0].ScriptTokenType);
    }


    [TestMethod]
    public void TestOneDoubleWrong()
    {
        LexerConfig conf = new LexerConfig();

        ScriptSplitter stringParser = new ScriptSplitter();
        string line = " 12.3A";
        bool res = stringParser.Split(1, line, conf.Separators, conf.StringSep, conf.CommentTag, out List<ScriptToken> listScriptTokens, out ScriptTokenType lastTokenType);

        // return false -> error
        Assert.IsFalse(res);
        Assert.AreEqual(1, listScriptTokens.Count);
        Assert.AreEqual(ScriptTokenType.WrongNumber, listScriptTokens[0].ScriptTokenType);
    }

    // 12 45
    [TestMethod]
    public void Test12_45_Ok()
    {
        LexerConfig conf = new LexerConfig();

        ScriptSplitter stringParser = new ScriptSplitter();
        string line = "12 45";
        bool res = stringParser.Split(1, line, conf.Separators, conf.StringSep, conf.CommentTag, out List<ScriptToken> listScriptTokens, out ScriptTokenType lastTokenType);

        Assert.IsTrue(res);
        Assert.AreEqual(2, listScriptTokens.Count);
        Assert.AreEqual(ScriptTokenType.Integer, listScriptTokens[0].ScriptTokenType);
        Assert.AreEqual(12, listScriptTokens[0].ValueInt);
        Assert.AreEqual(45, listScriptTokens[1].ValueInt);
    }

    // 12 45
    [TestMethod]
    public void Test_Spc_12_45_Ok()
    {
        LexerConfig conf = new LexerConfig();

        ScriptSplitter stringParser = new ScriptSplitter();
        string line = " 12 45";
        bool res = stringParser.Split(1, line, conf.Separators, conf.StringSep, conf.CommentTag, out List<ScriptToken> listScriptTokens, out ScriptTokenType lastTokenType);

        Assert.IsTrue(res);
        Assert.AreEqual(2, listScriptTokens.Count);
        Assert.AreEqual(ScriptTokenType.Integer, listScriptTokens[0].ScriptTokenType);
        Assert.AreEqual(12, listScriptTokens[0].ValueInt);
        Assert.AreEqual(45, listScriptTokens[1].ValueInt);
    }

    // 12 45
    [TestMethod]
    public void Test_Spc_12_45_Spc_Ok()
    {
        LexerConfig conf = new LexerConfig();

        ScriptSplitter stringParser = new ScriptSplitter();
        string line = " 12 45 ";
        bool res = stringParser.Split(1, line, conf.Separators, conf.StringSep, conf.CommentTag, out List<ScriptToken> listScriptTokens, out ScriptTokenType lastTokenType);

        Assert.IsTrue(res);
        Assert.AreEqual(2, listScriptTokens.Count);
        Assert.AreEqual(ScriptTokenType.Integer, listScriptTokens[0].ScriptTokenType);
        Assert.AreEqual(12, listScriptTokens[0].ValueInt);
        Assert.AreEqual(ScriptTokenType.Integer, listScriptTokens[1].ScriptTokenType);
        Assert.AreEqual(45, listScriptTokens[1].ValueInt);
    }


    // 12 - 23   -> 3 tok
    [TestMethod]
    public void Test_12_minus_23_Ok()
    {
        LexerConfig conf = new LexerConfig();

        ScriptSplitter stringParser = new ScriptSplitter();
        string line = "12-23";
        bool res = stringParser.Split(1, line, conf.Separators, conf.StringSep, conf.CommentTag, out List<ScriptToken> listScriptTokens, out ScriptTokenType lastTokenType);

        Assert.IsTrue(res);
        Assert.AreEqual(3, listScriptTokens.Count);
        Assert.AreEqual(ScriptTokenType.Integer, listScriptTokens[0].ScriptTokenType);
        Assert.AreEqual(12, listScriptTokens[0].ValueInt);

        Assert.AreEqual(ScriptTokenType.Separator, listScriptTokens[1].ScriptTokenType);
        Assert.AreEqual("-", listScriptTokens[1].Value);

        Assert.AreEqual(23, listScriptTokens[2].ValueInt);
    }

    // 12 - 23   -> 3 tok
    [TestMethod]
    public void Test_12_spc_minus_23_Ok()
    {
        LexerConfig conf = new LexerConfig();

        ScriptSplitter stringParser = new ScriptSplitter();
        string line = "12 -23";
        bool res = stringParser.Split(1, line, conf.Separators, conf.StringSep, conf.CommentTag, out List<ScriptToken> listScriptTokens, out ScriptTokenType lastTokenType);

        Assert.IsTrue(res);
        Assert.AreEqual(3, listScriptTokens.Count);
        Assert.AreEqual(ScriptTokenType.Integer, listScriptTokens[0].ScriptTokenType);
        Assert.AreEqual(12, listScriptTokens[0].ValueInt);

        Assert.AreEqual(ScriptTokenType.Separator, listScriptTokens[1].ScriptTokenType);
        Assert.AreEqual("-", listScriptTokens[1].Value);

        Assert.AreEqual(23, listScriptTokens[2].ValueInt);
    }

    // 12 - 23   -> 3 tok
    [TestMethod]
    public void Test_12_spc_minus_spc_23_Ok()
    {
        LexerConfig conf = new LexerConfig();

        ScriptSplitter stringParser = new ScriptSplitter();
        string line = "12 - 23";
        bool res = stringParser.Split(1, line, conf.Separators, conf.StringSep, conf.CommentTag, out List<ScriptToken> listScriptTokens, out ScriptTokenType lastTokenType);

        Assert.IsTrue(res);
        Assert.AreEqual(3, listScriptTokens.Count);
        Assert.AreEqual(ScriptTokenType.Integer, listScriptTokens[0].ScriptTokenType);
        Assert.AreEqual(12, listScriptTokens[0].ValueInt);

        Assert.AreEqual(ScriptTokenType.Separator, listScriptTokens[1].ScriptTokenType);
        Assert.AreEqual("-", listScriptTokens[1].Value);

        Assert.AreEqual(23, listScriptTokens[2].ValueInt);
    }

    // 12 + -23  -> 4 tok
    [TestMethod]
    public void Test_12_spc_plus_spc_minus_23_Ok()
    {
        LexerConfig conf = new LexerConfig();

        ScriptSplitter stringParser = new ScriptSplitter();
        string line = "12 + -23";
        bool res = stringParser.Split(1, line, conf.Separators, conf.StringSep, conf.CommentTag, out List<ScriptToken> listScriptTokens, out ScriptTokenType lastTokenType);

        Assert.IsTrue(res);
        Assert.AreEqual(4, listScriptTokens.Count);
        Assert.AreEqual(ScriptTokenType.Integer, listScriptTokens[0].ScriptTokenType);
        Assert.AreEqual(12, listScriptTokens[0].ValueInt);

        Assert.AreEqual(ScriptTokenType.Separator, listScriptTokens[1].ScriptTokenType);
        Assert.AreEqual("+", listScriptTokens[1].Value);

        Assert.AreEqual(ScriptTokenType.Separator, listScriptTokens[2].ScriptTokenType);
        Assert.AreEqual("-", listScriptTokens[2].Value);

        Assert.AreEqual(23, listScriptTokens[3].ValueInt);
    }

    // a>-12  -> 4 tok
    [TestMethod]
    public void Test_if_a_greater_minus_12_Ok()
    {
        LexerConfig conf = new LexerConfig();

        ScriptSplitter stringParser = new ScriptSplitter();
        string line = "a>-12";
        bool res = stringParser.Split(1, line, conf.Separators, conf.StringSep, conf.CommentTag, out List<ScriptToken> listScriptTokens, out ScriptTokenType lastTokenType);

        Assert.IsTrue(res);
        Assert.AreEqual(4, listScriptTokens.Count);
        Assert.AreEqual(ScriptTokenType.Name, listScriptTokens[0].ScriptTokenType);
        Assert.AreEqual("a", listScriptTokens[0].Value);

        Assert.AreEqual(ScriptTokenType.Separator, listScriptTokens[1].ScriptTokenType);
        Assert.AreEqual(">", listScriptTokens[1].Value);

        Assert.AreEqual(ScriptTokenType.Separator, listScriptTokens[2].ScriptTokenType);
        Assert.AreEqual("-", listScriptTokens[2].Value);

        Assert.AreEqual(ScriptTokenType.Integer, listScriptTokens[3].ScriptTokenType);
        Assert.AreEqual(12, listScriptTokens[3].ValueInt);
    }

    // "A.Cell "
    [TestMethod]
    public void Test_A_dot_Spc_Cell_spc_Ok()
    {
        LexerConfig conf = new LexerConfig();

        ScriptSplitter stringParser = new ScriptSplitter();
        string line = "A.Cell ";
        bool res = stringParser.Split(1, line, conf.Separators, conf.StringSep, conf.CommentTag, out List<ScriptToken> listScriptTokens, out ScriptTokenType lastTokenType);

        Assert.IsTrue(res);
        Assert.AreEqual(3, listScriptTokens.Count);
        Assert.AreEqual(ScriptTokenType.Name, listScriptTokens[0].ScriptTokenType);
        Assert.AreEqual("A", listScriptTokens[0].Value);
        Assert.AreEqual(ScriptTokenType.Separator, listScriptTokens[1].ScriptTokenType);
        Assert.AreEqual(".", listScriptTokens[1].Value);
        Assert.AreEqual(ScriptTokenType.Name, listScriptTokens[2].ScriptTokenType);
        Assert.AreEqual("Cell", listScriptTokens[2].Value);
    }

    // "A.Cell "
    [TestMethod]
    public void Test_A_dot_Spc_Cell_Ok()
    {
        LexerConfig conf = new LexerConfig();

        ScriptSplitter stringParser = new ScriptSplitter();
        string line = "A.Cell";
        bool res = stringParser.Split(1, line, conf.Separators, conf.StringSep, conf.CommentTag, out List<ScriptToken> listScriptTokens, out ScriptTokenType lastTokenType);

        Assert.IsTrue(res);
        Assert.AreEqual(3, listScriptTokens.Count);
        Assert.AreEqual(ScriptTokenType.Name, listScriptTokens[0].ScriptTokenType);
        Assert.AreEqual("A", listScriptTokens[0].Value);
        Assert.AreEqual(ScriptTokenType.Separator, listScriptTokens[1].ScriptTokenType);
        Assert.AreEqual(".", listScriptTokens[1].Value);
        Assert.AreEqual(ScriptTokenType.Name, listScriptTokens[2].ScriptTokenType);
        Assert.AreEqual("Cell", listScriptTokens[2].Value);
    }

    // "A."
    [TestMethod]
    public void Test_A_dot_Ok()
    {
        LexerConfig conf = new LexerConfig();

        ScriptSplitter stringParser = new ScriptSplitter();
        string line = "A.";
        bool res = stringParser.Split(1, line, conf.Separators, conf.StringSep, conf.CommentTag, out List<ScriptToken> listScriptTokens, out ScriptTokenType lastTokenType);

        Assert.IsTrue(res);
        Assert.AreEqual(2, listScriptTokens.Count);
        Assert.AreEqual(ScriptTokenType.Name, listScriptTokens[0].ScriptTokenType);
        Assert.AreEqual("A", listScriptTokens[0].Value);
        Assert.AreEqual(ScriptTokenType.Separator, listScriptTokens[1].ScriptTokenType);
        Assert.AreEqual(".", listScriptTokens[1].Value);
    }

    // "A. "
    [TestMethod]
    public void Test_A_dot_spc_Ok()
    {
        LexerConfig conf = new LexerConfig();

        ScriptSplitter stringParser = new ScriptSplitter();
        string line = "A. ";
        bool res = stringParser.Split(1, line, conf.Separators, conf.StringSep, conf.CommentTag, out List<ScriptToken> listScriptTokens, out ScriptTokenType lastTokenType);

        Assert.IsTrue(res);
        Assert.AreEqual(2, listScriptTokens.Count);
        Assert.AreEqual(ScriptTokenType.Name, listScriptTokens[0].ScriptTokenType);
        Assert.AreEqual("A", listScriptTokens[0].Value);
        Assert.AreEqual(ScriptTokenType.Separator, listScriptTokens[1].ScriptTokenType);
        Assert.AreEqual(".", listScriptTokens[1].Value);
    }

}
