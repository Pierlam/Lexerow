using FakeItEasy;
using Lexerow.Core.ScriptCompile.Parse;
using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;
using Lexerow.Core.System.ScriptDef;
using Lexerow.Core.Tests._05_Common;

namespace Lexerow.Core.Tests.ScriptParser;

/// <summary>
/// Test script lexical analyzer on SelectFiles instr.
/// </summary>
[TestClass]
public class ScriptParserSelectFilesTests
{
    /// <summary>
    /// file=SelectFiles("data.xslx")
    ///
    /// ----
    /// file; =; SelectFiles; (; "data.xlsx"; )
    ///
    /// ----
    /// The compilation return:
    ///  -SetVar:
    ///      Instrleft:  ObjectName: file
    ///      InstrRight: SelectFiles, p="data.xlsx"
    /// </summary>
    [TestMethod]
    public void FileEqSelectFilesOk()
    {
        int numLine = 0;
        List<ScriptLineTokens> scriptTokens = new List<ScriptLineTokens>();

        //-line #1
        TestTokensBuilder.AddLineSelectFiles(numLine++, scriptTokens, "file", "\"data.xlsx\"");

        //==>just to check the content of the script
        //var scriptCheck = TestTokens2ScriptBuilder.BuildScript(script);

        //==> Parse the script tokens
        Parser parser = new Parser(A.Fake<IActivityLogger>());
        Result result = new Result();
        var prog = TestInstrBuilder.CreateProgram();
        bool res = parser.Process(result, scriptTokens, prog);

        //==> Check the result
        Assert.IsTrue(res);
        Assert.AreEqual(1, prog.ListInstr.Count);

        // SetVar
        Assert.AreEqual(InstrType.SetVar, prog.ListInstr[0].InstrType);
        InstrSetVar instrSetVar = prog.ListInstr[0] as InstrSetVar;

        // InstrLeft: ObjectName
        InstrObjectName instrObjectName = instrSetVar.InstrLeft as InstrObjectName;
        Assert.IsNotNull(instrObjectName);
        Assert.AreEqual("file", instrObjectName.ObjectName);

        // InstrRight: SelectFiles
        InstrSelectFiles instrOpenExcel = instrSetVar.InstrRight as InstrSelectFiles;
        Assert.IsNotNull(instrOpenExcel);

        // OpenExcel Param
        Assert.AreEqual(1, instrOpenExcel.ListInstrParams.Count);
        InstrValue instrValue = instrOpenExcel.ListInstrParams[0] as InstrValue;
        Assert.IsNotNull(instrValue);
        Assert.AreEqual("data.xlsx", instrValue.RawValue);
        Assert.AreEqual("data.xlsx", (instrValue.ValueBase as ValueString).Val);
    }

    /// <summary>
    /// name="data.xslx"
    /// file=SelectFiles(name)
    ///
    /// ----
    /// name; =; "data.xlsx"
    /// file; =; SelectFiles; (; name; )
    ///
    /// ----
    /// The compilation return:
    ///  -SetVar:
    ///      Instrleft:  ObjectName: name
    ///      InstrRight: ConstValue: "data.xlsx"
    ///
    ///  -SetVar:
    ///      Instrleft:  ObjectName: file
    ///      InstrRight: SelectFiles, p=name
    /// </summary>
    [TestMethod]
    public void SetVarFileEqSelectFilesOk()
    {
        List<ScriptLineTokens> scriptTokens = new List<ScriptLineTokens>();
        ScriptLineTokens line;

        //-line #1
        line = new ScriptLineTokens();
        line.AddTokenName(1, 1, "name");
        line.AddTokenSeparator(1, 1, "=");
        line.AddTokenString(1, 1, "\"data.xlsx\"");
        scriptTokens.Add(line);

        //-build one line of tokens
        line = new ScriptLineTokens();
        line.AddTokenName(2, 1, "file");
        line.AddTokenSeparator(2, 1, "=");
        line.AddTokenName(2, 1, "SelectFiles");
        line.AddTokenSeparator(2, 1, "(");
        line.AddTokenName(2, 1, "name");
        line.AddTokenSeparator(2, 1, ")");
        scriptTokens.Add(line);

        //==>just to check the content of the script
        //var scriptCheck = TestTokens2ScriptBuilder.BuildScript(script);

        //==> Parse the script tokens
        Parser parser = new Parser(A.Fake<IActivityLogger>());
        Result result = new Result();
        var prog = TestInstrBuilder.CreateProgram();
        bool res = parser.Process(result, scriptTokens, prog);

        //==> Check the result
        Assert.IsTrue(res);
        Assert.AreEqual(2, prog.ListInstr.Count);

        //--SetVar #1
        Assert.AreEqual(InstrType.SetVar, prog.ListInstr[0].InstrType);
        InstrSetVar instrSetVar = prog.ListInstr[0] as InstrSetVar;

        // InstrLeft: ObjectName
        InstrObjectName instrObjectName = instrSetVar.InstrLeft as InstrObjectName;
        Assert.IsNotNull(instrObjectName);
        Assert.AreEqual("name", instrObjectName.ObjectName);

        // InstrRight: ConstValue
        InstrValue instrValue = instrSetVar.InstrRight as InstrValue;
        Assert.IsNotNull(instrValue);
        Assert.AreEqual("data.xlsx", instrValue.RawValue);
        Assert.AreEqual("data.xlsx", (instrValue.ValueBase as ValueString).Val);

        //--SetVar #2
        Assert.AreEqual(InstrType.SetVar, prog.ListInstr[1].InstrType);
        instrSetVar = prog.ListInstr[1] as InstrSetVar;

        // InstrLeft: ObjectName
        instrObjectName = instrSetVar.InstrLeft as InstrObjectName;
        Assert.IsNotNull(instrObjectName);
        Assert.AreEqual("file", instrObjectName.ObjectName);

        // InstrRight: SelectFiles
        var instrOpenExcel = instrSetVar.InstrRight as InstrSelectFiles;
        Assert.IsNotNull(instrOpenExcel);

        // OpenExcel Param -> object name
        Assert.AreEqual(1, instrOpenExcel.ListInstrParams.Count);
        instrObjectName = instrOpenExcel.ListInstrParams[0] as InstrObjectName;
        Assert.IsNotNull(instrObjectName);
        Assert.AreEqual("name", instrObjectName.ObjectName);
    }

    /// <summary>
    /// file=Qwerty()
    /// Error, Wrong/Not expected function name.
    /// </summary>
    [TestMethod]
    public void FileEqQwertyWrong()
    {
        //-build one line of tokens
        ScriptLineTokens line = new ScriptLineTokens();
        line.AddTokenName(1, 1, "file");
        line.AddTokenSeparator(1, 1, "=");
        line.AddTokenName(1, 1, "Qwerty");
        line.AddTokenSeparator(1, 1, "(");
        line.AddTokenString(1, 1, "\"data.xlsx\"");
        line.AddTokenSeparator(1, 1, ")");

        //-build source code lines of tokens
        List<ScriptLineTokens> scriptTokens = [line];

        //==>just to check the content of the script
        //var scriptCheck = TestTokens2ScriptBuilder.BuildScript(script);

        //==> Parse the script tokens
        Parser parser = new Parser(A.Fake<IActivityLogger>());
        Result result = new Result();
        var prog = TestInstrBuilder.CreateProgram();
        bool res = parser.Process(result, scriptTokens, prog);

        //==> Check the result

        Assert.IsFalse(res);
        Assert.AreEqual(0, prog.ListInstr.Count);
        Assert.AreEqual(1, result.ListError.Count);
        Assert.AreEqual(ErrorCode.ParserTokenNotExpected, result.ListError[0].ErrorCode);
    }

    /// <summary>
    /// "john"=SelectFiles(..)
    /// error -> variable expected
    /// </summary>
    [TestMethod]
    public void StrJohnEqSelectFilesWrong()
    {
        //-build one line of tokens
        ScriptLineTokens line = new ScriptLineTokens();
        line.AddTokenString(1, 1, "\"john\"");
        line.AddTokenSeparator(1, 1, "=");
        line.AddTokenName(1, 1, "SelectFiles");
        line.AddTokenSeparator(1, 1, "(");
        line.AddTokenString(1, 1, "\"data.xlsx\"");
        line.AddTokenSeparator(1, 1, ")");

        //-build source code lines of tokens
        List<ScriptLineTokens> scriptTokens = [line];

        //==>just to check the content of the script
        //var scriptCheck = TestTokens2ScriptBuilder.BuildScript(script);

        //==> Parse the script tokens
        Parser parser = new Parser(A.Fake<IActivityLogger>());
        Result result = new Result();
        var prog = TestInstrBuilder.CreateProgram();
        bool res = parser.Process(result, scriptTokens, prog);

        //==> Check the result
        Assert.IsFalse(res);
        Assert.AreEqual(0, prog.ListInstr.Count);
        Assert.AreEqual(1, result.ListError.Count);
        Assert.AreEqual(ErrorCode.ParserTokenNotExpected, result.ListError[0].ErrorCode);
    }

    /// <summary>
    /// 12=SelectFiles(...)
    /// error -> variable expected
    /// </summary>
    [TestMethod]
    public void i12EqSelectFilesWrong()
    {        
        //-build one line of tokens
        ScriptLineTokens  line = new ScriptLineTokens();
        line.AddTokenInteger(1, 1, 12);
        line.AddTokenSeparator(1, 1, "=");
        line.AddTokenName(1, 1, "SelectFiles");
        line.AddTokenSeparator(1, 1, "(");
        line.AddTokenString(1, 1, "\"data.xlsx\"");
        line.AddTokenSeparator(1, 1, ")");

        //-build source code lines of tokens
        List<ScriptLineTokens> scriptTokens = [line];

        //==>just to check the content of the script
        //var scriptCheck = TestTokens2ScriptBuilder.BuildScript(script);

        //==> Parse the script tokens
        Parser parser = new Parser(A.Fake<IActivityLogger>());
        Result result = new Result();
        var prog = TestInstrBuilder.CreateProgram();
        bool res = parser.Process(result, scriptTokens, prog);

        //==> Check the result
        Assert.IsFalse(res);
        Assert.AreEqual(0, prog.ListInstr.Count);
        Assert.AreEqual(1, result.ListError.Count);
        Assert.AreEqual(ErrorCode.ParserTokenNotExpected, result.ListError[0].ErrorCode);
    }

    /// <summary>
    /// file = SelectFiles()
    /// error -> param is missing
    /// </summary>
    [TestMethod]
    public void FileEqSelectFilesParamMissingWrong()
    {
        //-build one line of tokens
        ScriptLineTokens line = new ScriptLineTokens();
        line.AddTokenName(1, 1, "file");
        line.AddTokenSeparator(1, 1, "=");
        line.AddTokenName(1, 1, "SelectFiles");
        line.AddTokenSeparator(1, 1, "(");
        line.AddTokenSeparator(1, 1, ")");

        //-build source code lines of tokens
        List<ScriptLineTokens> scriptTokens = [line];

        //==>just to check the content of the script
        //var scriptCheck = TestTokens2ScriptBuilder.BuildScript(script);

        //==> Parse the script tokens
        Parser parser = new Parser(A.Fake<IActivityLogger>());
        Result result = new Result();
        var prog = TestInstrBuilder.CreateProgram();
        bool res = parser.Process(result, scriptTokens, prog);

        //==> Check the result
        Assert.IsFalse(res);
        Assert.AreEqual(0, prog.ListInstr.Count);
        Assert.AreEqual(1, result.ListError.Count);
        Assert.AreEqual(ErrorCode.ParserFctParamCountWrong, result.ListError[0].ErrorCode);
    }

    /// <summary>
    /// file = SelectFiles(12)
    /// error -> param type is wrong
    /// </summary>
    [TestMethod]
    public void FileEqSelectFilesParam12Wrong()
    {
        //-build one line of tokens
        ScriptLineTokens line = new ScriptLineTokens();
        line.AddTokenName(1, 1, "file");
        line.AddTokenSeparator(1, 1, "=");
        line.AddTokenName(1, 1, "SelectFiles");
        line.AddTokenSeparator(1, 1, "(");
        line.AddTokenInteger(1, 1, 12);
        line.AddTokenSeparator(1, 1, ")");

        //-build source code lines of tokens
        List<ScriptLineTokens> scriptTokens = [line];

        //==>just to check the content of the script
        //var scriptCheck = TestTokens2ScriptBuilder.BuildScript(script);

        //==> Parse the script tokens
        Parser parser = new Parser(A.Fake<IActivityLogger>());
        Result result = new Result();
        var prog = TestInstrBuilder.CreateProgram();
        bool res = parser.Process(result, scriptTokens, prog);

        //==> Check the result
        Assert.IsFalse(res);
        Assert.AreEqual(0, prog.ListInstr.Count);
        Assert.AreEqual(1, result.ListError.Count);
        Assert.AreEqual(ErrorCode.ParserFctParamTypeWrong, result.ListError[0].ErrorCode);
    }

    /// <summary>
    /// file = SelectFiles(f)
    /// error -> param f is not defined before
    /// </summary>
    [TestMethod]
    public void FileEqSelectFilesParamvarFWrong()
    {
        //-build one line of tokens
        ScriptLineTokens line = new ScriptLineTokens();
        line.AddTokenName(1, 1, "file");
        line.AddTokenSeparator(1, 1, "=");
        line.AddTokenName(1, 1, "SelectFiles");
        line.AddTokenSeparator(1, 1, "(");
        line.AddTokenName(1, 1, "f");
        line.AddTokenSeparator(1, 1, ")");

        //-build source code lines of tokens
        List<ScriptLineTokens> scriptTokens = [line];

        //==>just to check the content of the script
        //var scriptCheck = TestTokens2ScriptBuilder.BuildScript(script);

        //==> Parse the script tokens
        Parser parser = new Parser(A.Fake<IActivityLogger>());
        Result result = new Result();
        var prog = TestInstrBuilder.CreateProgram();
        bool res = parser.Process(result, scriptTokens, prog);

        //==> Check the result
        Assert.IsFalse(res);
        Assert.AreEqual(0, prog.ListInstr.Count);
        Assert.AreEqual(1, result.ListError.Count);
        Assert.AreEqual(ErrorCode.ParserFctParamVarNotDefined, result.ListError[0].ErrorCode);
    }

    /// <summary>
    /// file=SelectFiles
    /// error -> open-closed bracket missing (param is missing also)
    /// </summary>
    [TestMethod]
    public void FileEqSelectFilesNoParamWrong()
    {
        //-build one line of tokens
        ScriptLineTokens line = new ScriptLineTokens();
        line.AddTokenName(1, 1, "file");
        line.AddTokenSeparator(1, 1, "=");
        line.AddTokenName(1, 1, "SelectFiles");

        //-build source code lines of tokens
        List<ScriptLineTokens> scriptTokens = [line];

        //==>just to check the content of the script
        //var scriptCheck = TestTokens2ScriptBuilder.BuildScript(script);

        //==> Parse the script tokens
        Parser parser = new Parser(A.Fake<IActivityLogger>());
        Result result = new Result();
        var prog = TestInstrBuilder.CreateProgram();
        bool res = parser.Process(result, scriptTokens, prog);

        //==> Check the result
        Assert.IsFalse(res);
        Assert.AreEqual(0, prog.ListInstr.Count);
        Assert.AreEqual(1, result.ListError.Count);
        Assert.AreEqual(ErrorCode.ParserFctParamCountWrong, result.ListError[0].ErrorCode);
    }

    /// <summary>
    /// then = SelectFiles("dd")
    /// error -> nom variable non-authorisé, mot-clé réservé

    /// </summary>
    [TestMethod]
    public void ThenEqSelectFilesWrong()
    {
        //-build one line of tokens
        ScriptLineTokens line = new ScriptLineTokens();
        line.AddTokenName(1, 1, "then");
        line.AddTokenSeparator(1, 1, "=");
        line.AddTokenName(1, 1, "SelectFiles");
        line.AddTokenSeparator(1, 1, "(");
        line.AddTokenString(1, 1, "\"data.xlsx\"");
        line.AddTokenSeparator(1, 1, ")");

        //-build source code lines of tokens
        List<ScriptLineTokens> scriptTokens = [line];

        //==>just to check the content of the script
        //var scriptCheck = TestTokens2ScriptBuilder.BuildScript(script);

        //==> Parse the script tokens
        Parser parser = new Parser(A.Fake<IActivityLogger>());
        Result result = new Result();
        var prog = TestInstrBuilder.CreateProgram();
        bool res = parser.Process(result, scriptTokens, prog);

        //==> Check the result
        Assert.IsFalse(res);
        Assert.AreEqual(0, prog.ListInstr.Count);
        Assert.AreEqual(1, result.ListError.Count);
        Assert.AreEqual(ErrorCode.ParserTokenNotExpected, result.ListError[0].ErrorCode);
        Assert.AreEqual("then", result.ListError[0].Param);
    }

    /// <summary>
    /// file = SelectFiles("file.xlsx") load
    /// error -> Object name load not expected
    /// </summary>
    [TestMethod]
    public void ThenEqSelectFilesLoadWrong()
    {
        //-build one line of tokens
        ScriptLineTokens line = new ScriptLineTokens();
        line.AddTokenName(1, 1, "file");
        line.AddTokenSeparator(1, 1, "=");
        line.AddTokenName(1, 1, "SelectFiles");
        line.AddTokenSeparator(1, 1, "(");
        line.AddTokenString(1, 1, "\"data.xlsx\"");
        line.AddTokenSeparator(1, 1, ")");
        line.AddTokenName(1, 1, "load");

        //-build source code lines of tokens
        List<ScriptLineTokens> scriptTokens = [line];

        //==>just to check the content of the script
        //var scriptCheck = TestTokens2ScriptBuilder.BuildScript(script);

        //==> Parse the script tokens
        Parser parser = new Parser(A.Fake<IActivityLogger>());
        Result result = new Result();
        var prog = TestInstrBuilder.CreateProgram();
        bool res = parser.Process(result, scriptTokens, prog);

        //==> Check the result
        Assert.IsFalse(res);
        Assert.AreEqual(0, prog.ListInstr.Count);
        Assert.AreEqual(1, result.ListError.Count);

        Assert.AreEqual(ErrorCode.ParserTokenNotExpected, result.ListError[0].ErrorCode);
    }

    /// <summary>
    /// SelectFiles("data.xlsx")
    /// error -> fct result not used

    /// </summary>
    [TestMethod]
    public void SelectFilesWrong()
    {
        //-build one line of tokens
        ScriptLineTokens line = new ScriptLineTokens();
        line.AddTokenName(1, 1, "SelectFiles");
        line.AddTokenSeparator(1, 1, "(");
        line.AddTokenString(1, 1, "\"data.xlsx\"");
        line.AddTokenSeparator(1, 1, ")");

        //-build source code lines of tokens
        List<ScriptLineTokens> scriptTokens = [line];

        //==>just to check the content of the script
        //var scriptCheck = TestTokens2ScriptBuilder.BuildScript(script);

        //==> Parse the script tokens
        Parser parser = new Parser(A.Fake<IActivityLogger>());
        Result result = new Result();
        var prog = TestInstrBuilder.CreateProgram();
        bool res = parser.Process(result, scriptTokens, prog);

        //==> Check the result
        Assert.IsFalse(res);
        Assert.AreEqual(0, prog.ListInstr.Count);
        Assert.AreEqual(1, result.ListError.Count);

        Assert.AreEqual(ErrorCode.ParserFctResultNotSet, result.ListError[0].ErrorCode);
    }

    // file=
    //   OpenExcel("dd")    on 2 lines, allowed?
}