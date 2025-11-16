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
        List<ScriptLineTokens> script = new List<ScriptLineTokens>();
        ScriptLineTokens line;

        //-line #1
        line = TestTokensBuilder.BuildSelectFiles("file", "\"data.xlsx\"");
        script.Add(line);

        var logger = A.Fake<IActivityLogger>();
        Parser sa = new Parser(logger);

        ExecResult execResult = new ExecResult();
        bool res = sa.Process(execResult, script, out List<InstrBase> listInstr);

        Assert.IsTrue(res);
        Assert.AreEqual(1, listInstr.Count);

        // SetVar
        Assert.AreEqual(InstrType.SetVar, listInstr[0].InstrType);
        InstrSetVar instrSetVar = listInstr[0] as InstrSetVar;

        // InstrLeft: ObjectName
        InstrObjectName instrObjectName = instrSetVar.InstrLeft as InstrObjectName;
        Assert.IsNotNull(instrObjectName);
        Assert.AreEqual("file", instrObjectName.ObjectName);

        // InstrRight: SelectFiles
        InstrSelectFiles instrOpenExcel = instrSetVar.InstrRight as InstrSelectFiles;
        Assert.IsNotNull(instrOpenExcel);

        // OpenExcel Param
        Assert.AreEqual(1, instrOpenExcel.ListInstrParams.Count);
        InstrConstValue instrConstValue = instrOpenExcel.ListInstrParams[0] as InstrConstValue;
        Assert.IsNotNull(instrConstValue);
        Assert.AreEqual("data.xlsx", instrConstValue.RawValue);
        Assert.AreEqual("data.xlsx", (instrConstValue.ValueBase as ValueString).Val);
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
        List<ScriptLineTokens> script = new List<ScriptLineTokens>();
        ScriptLineTokens line;

        //-line #1
        line = new ScriptLineTokens();
        line.AddTokenName(1, 1, "name");
        line.AddTokenSeparator(1, 1, "=");
        line.AddTokenString(1, 1, "\"data.xlsx\"");
        script.Add(line);

        //-build one line of tokens
        line = new ScriptLineTokens();
        line.AddTokenName(2, 1, "file");
        line.AddTokenSeparator(2, 1, "=");
        line.AddTokenName(2, 1, "SelectFiles");
        line.AddTokenSeparator(2, 1, "(");
        line.AddTokenName(2, 1, "name");
        line.AddTokenSeparator(2, 1, ")");
        script.Add(line);

        var logger = A.Fake<IActivityLogger>();
        Parser sa = new Parser(logger);

        ExecResult execResult = new ExecResult();
        bool res = sa.Process(execResult, script, out List<InstrBase> listInstr);

        Assert.IsTrue(res);
        Assert.AreEqual(2, listInstr.Count);

        //--SetVar #1
        Assert.AreEqual(InstrType.SetVar, listInstr[0].InstrType);
        InstrSetVar instrSetVar = listInstr[0] as InstrSetVar;

        // InstrLeft: ObjectName
        InstrObjectName instrObjectName = instrSetVar.InstrLeft as InstrObjectName;
        Assert.IsNotNull(instrObjectName);
        Assert.AreEqual("name", instrObjectName.ObjectName);

        // InstrRight: ConstValue
        InstrConstValue instrConstValue = instrSetVar.InstrRight as InstrConstValue;
        Assert.IsNotNull(instrConstValue);
        Assert.AreEqual("data.xlsx", instrConstValue.RawValue);
        Assert.AreEqual("data.xlsx", (instrConstValue.ValueBase as ValueString).Val);

        //--SetVar #2
        Assert.AreEqual(InstrType.SetVar, listInstr[1].InstrType);
        instrSetVar = listInstr[1] as InstrSetVar;

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
        ScriptLineTokens slt;

        //-build one line of tokens
        slt = new ScriptLineTokens();
        slt.AddTokenName(1, 1, "file");
        slt.AddTokenSeparator(1, 1, "=");
        slt.AddTokenName(1, 1, "Qwerty");
        slt.AddTokenSeparator(1, 1, "(");
        slt.AddTokenString(1, 1, "\"data.xlsx\"");
        slt.AddTokenSeparator(1, 1, ")");

        //-build source code lines of tokens
        List<ScriptLineTokens> lt = [slt];

        var logger = A.Fake<IActivityLogger>();
        Parser sa = new Parser(logger);

        ExecResult execResult = new ExecResult();
        bool res = sa.Process(execResult, lt, out List<InstrBase> listInstr);

        Assert.IsFalse(res);
        Assert.AreEqual(0, listInstr.Count);
        Assert.AreEqual(1, execResult.ListError.Count);
        Assert.AreEqual(ErrorCode.ParserTokenNotExpected, execResult.ListError[0].ErrorCode);
    }

    /// <summary>
    /// "john"=SelectFiles(..)
    /// error -> variable expected
    /// </summary>
    [TestMethod]
    public void StrJohnEqSelectFilesWrong()
    {
        ScriptLineTokens slt;

        //-build one line of tokens
        slt = new ScriptLineTokens();
        slt.AddTokenString(1, 1, "\"john\"");
        slt.AddTokenSeparator(1, 1, "=");
        slt.AddTokenName(1, 1, "SelectFiles");
        slt.AddTokenSeparator(1, 1, "(");
        slt.AddTokenString(1, 1, "\"data.xlsx\"");
        slt.AddTokenSeparator(1, 1, ")");

        //-build source code lines of tokens
        List<ScriptLineTokens> lt = [slt];

        var logger = A.Fake<IActivityLogger>();
        Parser sa = new Parser(logger);

        ExecResult execResult = new ExecResult();
        bool res = sa.Process(execResult, lt, out List<InstrBase> listInstr);

        Assert.IsFalse(res);
        Assert.AreEqual(0, listInstr.Count);
        Assert.AreEqual(1, execResult.ListError.Count);
        Assert.AreEqual(ErrorCode.ParserTokenNotExpected, execResult.ListError[0].ErrorCode);
    }

    /// <summary>
    /// 12=SelectFiles(...)
    /// error -> variable expected
    /// </summary>
    [TestMethod]
    public void i12EqSelectFilesWrong()
    {
        ScriptLineTokens slt;

        //-build one line of tokens
        slt = new ScriptLineTokens();
        slt.AddTokenInteger(1, 1, 12);
        slt.AddTokenSeparator(1, 1, "=");
        slt.AddTokenName(1, 1, "SelectFiles");
        slt.AddTokenSeparator(1, 1, "(");
        slt.AddTokenString(1, 1, "\"data.xlsx\"");
        slt.AddTokenSeparator(1, 1, ")");

        //-build source code lines of tokens
        List<ScriptLineTokens> lt = [slt];

        var logger = A.Fake<IActivityLogger>();
        Parser sa = new Parser(logger);

        ExecResult execResult = new ExecResult();
        bool res = sa.Process(execResult, lt, out List<InstrBase> listInstr);

        Assert.IsFalse(res);
        Assert.AreEqual(0, listInstr.Count);
        Assert.AreEqual(1, execResult.ListError.Count);
        Assert.AreEqual(ErrorCode.ParserTokenNotExpected, execResult.ListError[0].ErrorCode);
    }

    /// <summary>
    /// file = SelectFiles()
    /// error -> param is missing
    /// </summary>
    [TestMethod]
    public void FileEqSelectFilesParamMissingWrong()
    {
        ScriptLineTokens slt;

        //-build one line of tokens
        slt = new ScriptLineTokens();
        slt.AddTokenName(1, 1, "file");
        slt.AddTokenSeparator(1, 1, "=");
        slt.AddTokenName(1, 1, "SelectFiles");
        slt.AddTokenSeparator(1, 1, "(");
        slt.AddTokenSeparator(1, 1, ")");

        //-build source code lines of tokens
        List<ScriptLineTokens> lt = [slt];

        var logger = A.Fake<IActivityLogger>();
        Parser sa = new Parser(logger);

        ExecResult execResult = new ExecResult();
        bool res = sa.Process(execResult, lt, out List<InstrBase> listInstr);

        Assert.IsFalse(res);
        Assert.AreEqual(0, listInstr.Count);
        Assert.AreEqual(1, execResult.ListError.Count);
        Assert.AreEqual(ErrorCode.ParserFctParamCountWrong, execResult.ListError[0].ErrorCode);
    }

    /// <summary>
    /// file = SelectFiles(12)
    /// error -> param type is wrong
    /// </summary>
    [TestMethod]
    public void FileEqSelectFilesParam12Wrong()
    {
        ScriptLineTokens slt;

        //-build one line of tokens
        slt = new ScriptLineTokens();
        slt.AddTokenName(1, 1, "file");
        slt.AddTokenSeparator(1, 1, "=");
        slt.AddTokenName(1, 1, "SelectFiles");
        slt.AddTokenSeparator(1, 1, "(");
        slt.AddTokenInteger(1, 1, 12);
        slt.AddTokenSeparator(1, 1, ")");

        //-build source code lines of tokens
        List<ScriptLineTokens> lt = [slt];

        var logger = A.Fake<IActivityLogger>();
        Parser sa = new Parser(logger);

        ExecResult execResult = new ExecResult();
        bool res = sa.Process(execResult, lt, out List<InstrBase> listInstr);

        Assert.IsFalse(res);
        Assert.AreEqual(0, listInstr.Count);
        Assert.AreEqual(1, execResult.ListError.Count);
        Assert.AreEqual(ErrorCode.ParserFctParamTypeWrong, execResult.ListError[0].ErrorCode);
    }

    /// <summary>
    /// file = SelectFiles(f)
    /// error -> param f is not defined before
    /// </summary>
    [TestMethod]
    public void FileEqSelectFilesParamvarFWrong()
    {
        ScriptLineTokens slt;

        //-build one line of tokens
        slt = new ScriptLineTokens();
        slt.AddTokenName(1, 1, "file");
        slt.AddTokenSeparator(1, 1, "=");
        slt.AddTokenName(1, 1, "SelectFiles");
        slt.AddTokenSeparator(1, 1, "(");
        slt.AddTokenName(1, 1, "f");
        slt.AddTokenSeparator(1, 1, ")");

        //-build source code lines of tokens
        List<ScriptLineTokens> lt = [slt];

        var logger = A.Fake<IActivityLogger>();
        Parser sa = new Parser(logger);

        ExecResult execResult = new ExecResult();
        bool res = sa.Process(execResult, lt, out List<InstrBase> listInstr);

        Assert.IsFalse(res);
        Assert.AreEqual(0, listInstr.Count);
        Assert.AreEqual(1, execResult.ListError.Count);
        Assert.AreEqual(ErrorCode.ParserFctParamVarNotDefined, execResult.ListError[0].ErrorCode);
    }

    /// <summary>
    /// file=SelectFiles
    /// error -> open-closed bracket missing (param is missing also)
    /// </summary>
    [TestMethod]
    public void FileEqSelectFilesNoParamWrong()
    {
        ScriptLineTokens slt;

        //-build one line of tokens
        slt = new ScriptLineTokens();
        slt.AddTokenName(1, 1, "file");
        slt.AddTokenSeparator(1, 1, "=");
        slt.AddTokenName(1, 1, "SelectFiles");

        //-build source code lines of tokens
        List<ScriptLineTokens> lt = [slt];

        var logger = A.Fake<IActivityLogger>();
        Parser sa = new Parser(logger);

        ExecResult execResult = new ExecResult();
        bool res = sa.Process(execResult, lt, out List<InstrBase> listInstr);

        Assert.IsFalse(res);
        Assert.AreEqual(0, listInstr.Count);
        Assert.AreEqual(1, execResult.ListError.Count);
        Assert.AreEqual(ErrorCode.ParserFctParamCountWrong, execResult.ListError[0].ErrorCode);
    }

    /// <summary>
    /// then = SelectFiles("dd")
    /// error -> nom variable non-authorisé, mot-clé réservé

    /// </summary>
    [TestMethod]
    public void ThenEqSelectFilesWrong()
    {
        ScriptLineTokens slt;

        //-build one line of tokens
        slt = new ScriptLineTokens();
        slt.AddTokenName(1, 1, "then");
        slt.AddTokenSeparator(1, 1, "=");
        slt.AddTokenName(1, 1, "SelectFiles");
        slt.AddTokenSeparator(1, 1, "(");
        slt.AddTokenString(1, 1, "\"data.xlsx\"");
        slt.AddTokenSeparator(1, 1, ")");

        //-build source code lines of tokens
        List<ScriptLineTokens> lt = [slt];

        var logger = A.Fake<IActivityLogger>();
        Parser sa = new Parser(logger);

        ExecResult execResult = new ExecResult();
        bool res = sa.Process(execResult, lt, out List<InstrBase> listInstr);

        Assert.IsFalse(res);
        Assert.AreEqual(0, listInstr.Count);
        Assert.AreEqual(1, execResult.ListError.Count);
        Assert.AreEqual(ErrorCode.ParserTokenNotExpected, execResult.ListError[0].ErrorCode);
        Assert.AreEqual("then", execResult.ListError[0].Param);
    }

    /// <summary>
    /// file = SelectFiles("file.xlsx") load
    /// error -> Object name load not expected
    /// </summary>
    [TestMethod]
    public void ThenEqSelectFilesLoadWrong()
    {
        ScriptLineTokens slt;

        //-build one line of tokens
        slt = new ScriptLineTokens();
        slt.AddTokenName(1, 1, "file");
        slt.AddTokenSeparator(1, 1, "=");
        slt.AddTokenName(1, 1, "SelectFiles");
        slt.AddTokenSeparator(1, 1, "(");
        slt.AddTokenString(1, 1, "\"data.xlsx\"");
        slt.AddTokenSeparator(1, 1, ")");
        slt.AddTokenName(1, 1, "load");

        //-build source code lines of tokens
        List<ScriptLineTokens> lt = [slt];

        var logger = A.Fake<IActivityLogger>();
        Parser sa = new Parser(logger);

        ExecResult execResult = new ExecResult();
        bool res = sa.Process(execResult, lt, out List<InstrBase> listInstr);

        Assert.IsFalse(res);
        Assert.AreEqual(0, listInstr.Count);
        Assert.AreEqual(1, execResult.ListError.Count);

        Assert.AreEqual(ErrorCode.ParserTokenNotExpected, execResult.ListError[0].ErrorCode);
    }

    /// <summary>
    /// SelectFiles("data.xlsx")
    /// error -> fct result not used

    /// </summary>
    [TestMethod]
    public void SelectFilesWrong()
    {
        ScriptLineTokens slt;

        //-build one line of tokens
        slt = new ScriptLineTokens();
        slt.AddTokenName(1, 1, "SelectFiles");
        slt.AddTokenSeparator(1, 1, "(");
        slt.AddTokenString(1, 1, "\"data.xlsx\"");
        slt.AddTokenSeparator(1, 1, ")");

        //-build source code lines of tokens
        List<ScriptLineTokens> lt = [slt];

        var logger = A.Fake<IActivityLogger>();
        Parser sa = new Parser(logger);

        ExecResult execResult = new ExecResult();
        bool res = sa.Process(execResult, lt, out List<InstrBase> listInstr);

        Assert.IsFalse(res);
        Assert.AreEqual(0, listInstr.Count);
        Assert.AreEqual(1, execResult.ListError.Count);

        Assert.AreEqual(ErrorCode.ParserFctResultNotSet, execResult.ListError[0].ErrorCode);
    }

    // file=
    //   OpenExcel("dd")    on 2 lines, allowed?
}