using FakeItEasy;
using Lexerow.Core.ScriptCompile.Parse;
using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;
using Lexerow.Core.System.ScriptDef;
using Lexerow.Core.Tests._05_Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Tests.ScriptParser;

/// <summary>
/// Test script parser on OnExcel instr.
/// Focus on FirstRow instr.
/// </summary>
[TestClass]
public class ScriptParserOnExcelFirstRowOkTests
{
    /// <summary>
    /// Default: sheet=0
    ///
    ///	OnExcel "file.xlsx"
    ///	  FirstRow 3
    ///   ForEach Row
    ///     If A.Cell>10 Then A.Cell=10
    ///   Next
    /// End OnExcel
    /// </summary>
    [TestMethod]
    public void OnExcelFirstRowValueOk()
    {
        int numLine = 1;
        List<ScriptLineTokens> scriptTokens = new List<ScriptLineTokens>();

        // OnExcel "data.xlsx"
        TestTokensBuilder.AddLineOnExcelFileString(numLine++, scriptTokens, "\"data.xlsx\"");

        // FirstRow 3
        TestTokensBuilder.AddLineFirstRow(numLine++, scriptTokens, 3);

        // ForEach Row
        TestTokensBuilder.AddLineForEachRow(numLine++, scriptTokens);

        // If A.Cell <= 10 Then A.Cell=10
        TestTokensBuilder.BuidIfColCellCompIntThenSetColCellInt(numLine++, scriptTokens, "A", "<=", 10, "A", 10);

        // Next
        TestTokensBuilder.AddLineNext(numLine++, scriptTokens);

        // End OnExcel
        TestTokensBuilder.AddLineEndOnExcel(numLine++, scriptTokens);

        //==>just to check the content of the script
        //var scriptCheck = TestTokens2ScriptBuilder.BuildScript(scriptTokens);

        //==> Parse the script tokens
        Parser parser = new Parser(A.Fake<IActivityLogger>());
        ExecResult execResult = new ExecResult();
        var prog = TestInstrBuilder.CreateProgram();
        bool res = parser.Process(execResult, scriptTokens, prog);

        //==> check the result
        Assert.IsTrue(res);
        Assert.AreEqual(1, prog.ListInstr.Count);

        // OnExcel
        Assert.AreEqual(InstrType.OnExcel, prog.ListInstr[0].InstrType);
        InstrOnExcel instrOnExcel = prog.ListInstr[0] as InstrOnExcel;

        // OnSheet
        Assert.AreEqual(1, instrOnExcel.ListSheets.Count);
        InstrOnSheet instrOnSheet = instrOnExcel.ListSheets[0];

        // one IfThen in OnSheet
        Assert.AreEqual(1, instrOnSheet.ListInstrForEachRow.Count);

        // check the firstrow value
        InstrValue instrValue = instrOnExcel.ListSheets[0].InstrFirstDataRow as InstrValue;
        Assert.IsNotNull(instrValue);
        Assert.AreEqual(3, ((instrValue.ValueBase) as ValueInt).Val);
    }

    /// <summary>
    /// Default: sheet=0
    ///
    ///	OnExcel "file.xlsx"
    ///	  FirstRow 0
    ///   ForEach Row
    ///     If A.Cell>10 Then A.Cell=10
    ///   Next
    /// End OnExcel
    /// </summary>
    [TestMethod]
    public void OnExcelFirstRowValueError()
    {
        int numLine = 1;
        List<ScriptLineTokens> scriptTokens = new List<ScriptLineTokens>();

        // OnExcel "data.xlsx"
        TestTokensBuilder.AddLineOnExcelFileString(numLine++, scriptTokens, "\"data.xlsx\"");

        // FirstRow 3
        TestTokensBuilder.AddLineFirstRow(numLine++, scriptTokens, 0);

        // ForEach Row
        TestTokensBuilder.AddLineForEachRow(numLine++, scriptTokens);

        // If A.Cell <= 10 Then A.Cell=10
        TestTokensBuilder.BuidIfColCellCompIntThenSetColCellInt(numLine++, scriptTokens, "A", "<=", 10, "A", 10);

        // Next
        TestTokensBuilder.AddLineNext(numLine++, scriptTokens);

        // End OnExcel
        TestTokensBuilder.AddLineEndOnExcel(numLine++, scriptTokens);

        //==>just to check the content of the script
        var scriptCheck = TestTokens2ScriptBuilder.BuildScript(scriptTokens);

        //==> Parse the script tokens
        Parser parser = new Parser(A.Fake<IActivityLogger>());
        ExecResult execResult = new ExecResult();
        var prog = TestInstrBuilder.CreateProgram();
        bool res = parser.Process(execResult, scriptTokens, prog);

        //==> check the result
        Assert.IsFalse(res);
        Assert.AreEqual(ErrorCode.ParserConstValueIntWrong, execResult.ListError[0].ErrorCode);
    }

    /// <summary>
    /// Default: sheet=0
    ///
    ///	OnExcel "file.xlsx"
    ///	  FirstRow 0
    ///   ForEach Row
    ///     If A.Cell>10 Then A.Cell=10
    ///   Next
    /// End OnExcel
    /// </summary>
    [TestMethod]
    public void OnExcelFirstRowZeroError()
    {
        int numLine = 1;
        List<ScriptLineTokens> scriptTokens = new List<ScriptLineTokens>();

        // OnExcel "data.xlsx"
        TestTokensBuilder.AddLineOnExcelFileString(numLine++, scriptTokens, "\"data.xlsx\"");

        // FirstRow 0
        TestTokensBuilder.AddLineFirstRow(numLine++, scriptTokens, 0);

        // ForEach Row
        TestTokensBuilder.AddLineForEachRow(numLine++, scriptTokens);

        // If A.Cell>10 Then A.Cell=10
        TestTokensBuilder.BuidIfColCellCompIntThenSetColCellInt(numLine++, scriptTokens, "A", ">", 10, "A", 10);

        // Next
        TestTokensBuilder.AddLineNext(numLine++, scriptTokens);

        // End OnExcel
        TestTokensBuilder.AddLineEndOnExcel(numLine++, scriptTokens);

        //==>just to check the content of the script
        var scriptCheck = TestTokens2ScriptBuilder.BuildScript(scriptTokens);

        //==> Parse the script tokens
        Parser parser = new Parser(A.Fake<IActivityLogger>());
        ExecResult execResult = new ExecResult();
        var prog = TestInstrBuilder.CreateProgram();
        bool res = parser.Process(execResult, scriptTokens, prog);

        //==> check the result
        Assert.IsFalse(res);
        Assert.AreEqual(ErrorCode.ParserConstValueIntWrong, execResult.ListError[0].ErrorCode);
    }

    /// <summary>
    /// Default: sheet=0
    ///
    /// r=3
    ///	OnExcel "file.xlsx"
    ///	  FirstRow r
    ///   ForEach Row
    ///     If A.Cell>10 Then A.Cell=10
    ///   Next
    /// End OnExcel
    /// </summary>
    [TestMethod]
    public void OnExcelFirstRowVarOk()
    {
        int numLine = 0;
        List<ScriptLineTokens> scriptTokens = new List<ScriptLineTokens>();

        // create var r=3
        TestTokensBuilder.AddLineSetVarInt(numLine++, scriptTokens, "r", 3);

        //OnExcel "file.xlsx"
        TestTokensBuilder.AddLineOnExcelFileString(numLine++, scriptTokens, "\"data.xlsx\"");

        // FirstRow r
        TestTokensBuilder.AddLineFirstRowVar(numLine++, scriptTokens, "r");

        // ForEach Row
        TestTokensBuilder.AddLineForEachRow(numLine++, scriptTokens);

        // If A.Cell>10 Then A.Cell=10
        TestTokensBuilder.BuidIfColCellCompIntThenSetColCellInt(numLine++, scriptTokens, "A", ">", 10, "A", 10);

        // Next
        TestTokensBuilder.AddLineNext(numLine++, scriptTokens);

        // End OnExcel
        TestTokensBuilder.AddLineEndOnExcel(numLine++, scriptTokens);

        //==>just to check the content of the script
        var scriptCheck = TestTokens2ScriptBuilder.BuildScript(scriptTokens);

        //==> Parse the script tokens
        Parser parser = new Parser(A.Fake<IActivityLogger>());
        ExecResult execResult = new ExecResult();
        var prog = TestInstrBuilder.CreateProgram();
        bool res = parser.Process(execResult, scriptTokens, prog);

        //==> check the result
        Assert.IsTrue(res);
        Assert.AreEqual(2, prog.ListInstr.Count);

        // SetVar r=3
        Assert.AreEqual(InstrType.SetVar, prog.ListInstr[0].InstrType);

        // OnExcel
        Assert.AreEqual(InstrType.OnExcel, prog.ListInstr[1].InstrType);
        InstrOnExcel instrOnExcel = prog.ListInstr[1] as InstrOnExcel;

        // OnSheet
        Assert.AreEqual(1, instrOnExcel.ListSheets.Count);
        InstrOnSheet instrOnSheet = instrOnExcel.ListSheets[0];

        // one IfThen in OnSheet
        Assert.AreEqual(1, instrOnSheet.ListInstrForEachRow.Count);

        // check the firstrow value, it's a var
        InstrObjectName instrObjectName = instrOnExcel.ListSheets[0].InstrFirstDataRow as InstrObjectName;
        Assert.IsNotNull(instrObjectName);
        Assert.AreEqual("r", instrObjectName.ObjectName);
    }

    /// <summary>
    /// Default: sheet=0
    ///
    ///	OnExcel "file.xlsx"
    ///	  FirstRow r
    ///   ForEach Row
    ///     If A.Cell>10 Then A.Cell=10
    ///   Next
    /// End OnExcel
    /// </summary>
    [TestMethod]
    public void OnExcelFirstRowVarNotExistsError()
    {
        int numLine = 0;
        List<ScriptLineTokens> scriptTokens = new List<ScriptLineTokens>();

        //OnExcel "file.xlsx"
        TestTokensBuilder.AddLineOnExcelFileString(numLine++, scriptTokens, "\"data.xlsx\"");

        // FirstRow r
        TestTokensBuilder.AddLineFirstRowVar(numLine++, scriptTokens, "r");

        // ForEach Row
        TestTokensBuilder.AddLineForEachRow(numLine++, scriptTokens);

        // If A.Cell>10 Then A.Cell=10
        TestTokensBuilder.BuidIfColCellCompIntThenSetColCellInt(numLine++, scriptTokens, "A", ">", 10, "A", 10);

        // Next
        TestTokensBuilder.AddLineNext(numLine++, scriptTokens);

        // End OnExcel
        TestTokensBuilder.AddLineEndOnExcel(numLine++, scriptTokens);

        //==>just to check the content of the script
        var scriptCheck = TestTokens2ScriptBuilder.BuildScript(scriptTokens);

        //==> Parse the script tokens
        Parser parser = new Parser(A.Fake<IActivityLogger>());
        ExecResult execResult = new ExecResult();
        var prog = TestInstrBuilder.CreateProgram();
        bool res = parser.Process(execResult, scriptTokens, prog);

        //==> check the result
        Assert.IsFalse(res);
        Assert.AreEqual(ErrorCode.ParserVarNotDefined, execResult.ListError[0].ErrorCode);
    }

    /// <summary>
    /// Default: sheet=0
    ///
    /// r=0                         ##should be > 0 !
    ///	OnExcel "file.xlsx"
    ///	  FirstRow r
    ///   ForEach Row
    ///     If A.Cell>10 Then A.Cell=10
    ///   Next
    /// End OnExcel
    /// </summary>
    [TestMethod]
    public void OnExcelFirstRowVarError()
    {
        int numLine = 0;
        List<ScriptLineTokens> scriptTokens = new List<ScriptLineTokens>();

        // create var r=3
        TestTokensBuilder.AddLineSetVarInt(numLine++, scriptTokens, "r", 0);

        //OnExcel "file.xlsx"
        TestTokensBuilder.AddLineOnExcelFileString(numLine++, scriptTokens, "\"data.xlsx\"");

        // FirstRow r
        TestTokensBuilder.AddLineFirstRowVar(numLine++, scriptTokens, "r");

        // ForEach Row
        TestTokensBuilder.AddLineForEachRow(numLine++, scriptTokens);

        // If A.Cell>10 Then A.Cell=10
        TestTokensBuilder.BuidIfColCellCompIntThenSetColCellInt(numLine++, scriptTokens, "A", ">", 10, "A", 10);

        // Next
        TestTokensBuilder.AddLineNext(numLine++, scriptTokens);

        // End OnExcel
        TestTokensBuilder.AddLineEndOnExcel(numLine++, scriptTokens);

        //==>just to check the content of the script
        var scriptCheck = TestTokens2ScriptBuilder.BuildScript(scriptTokens);

        //==> Parse the script tokens
        Parser parser = new Parser(A.Fake<IActivityLogger>());
        ExecResult execResult = new ExecResult();
        var prog = TestInstrBuilder.CreateProgram();
        bool res = parser.Process(execResult, scriptTokens, prog);

        //==> check the result
        Assert.IsFalse(res);
        Assert.AreEqual(ErrorCode.ParserConstValueIntWrong, execResult.ListError[0].ErrorCode);
    }
}