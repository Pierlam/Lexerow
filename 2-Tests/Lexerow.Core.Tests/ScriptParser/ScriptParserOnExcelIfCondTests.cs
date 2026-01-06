using Castle.Components.DictionaryAdapter.Xml;
using FakeItEasy;
using Lexerow.Core.ScriptCompile.Parse;
using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;
using Lexerow.Core.System.InstrDef;
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
/// Focus on If condition.
/// </summary>
[TestClass]
public class ScriptParserOnExcelIfCondTests
{
    /// <summary>
    /// Compile: OnExcel, very short version
    /// Result: one instruction OnExcel
    /// Implicite: sheet=0, FirstRow=1
    /// 
    /// a=true
    ///	OnExcel "file.xlsx"
    ///   ForEach Row
    ///     If a Then C.Cell=25
    ///   Next
    /// End OnExcel
    /// </summary>
    [TestMethod]
    public void varaEqTrueOnExcelIfaOk()
    {
        int numLine = 0;
        List<ScriptLineTokens> scriptTokens = new List<ScriptLineTokens>();

        // a=true
        TestTokensBuilder.AddLineSetVarStrBool(numLine++, scriptTokens, "a", "true");

        // OnExcel "data.xlsx"
        TestTokensBuilder.AddLineOnExcelFileString(numLine++, scriptTokens, "\"data.xlsx\"");

        // ForEach Row
        TestTokensBuilder.AddLineForEachRow(numLine++, scriptTokens);

        // If A.Cell >10 And B.Cell< 20 Then C.Cell=25  (in the same script line!!)
        var line = new ScriptLineTokens();
        line.AddTokenName(numLine++, 1, "If");
        line.AddTokenName(numLine, 1, "a");
        line.AddTokenName(numLine, 1, "Then");
        TestTokensBuilder.BuidColCellEqualInt(numLine++, line, "C", 25);
        scriptTokens.Add(line);

        // Next
        TestTokensBuilder.AddLineNext(numLine++, scriptTokens);

        // End OnExcel
        TestTokensBuilder.AddLineEndOnExcel(numLine++, scriptTokens);

        //==>just to check the content of the script
        var scriptCheck = TestTokens2ScriptBuilder.BuildScript(scriptTokens);

        //==> Parse the script tokens
        Parser parser = new Parser(A.Fake<IActivityLogger>());
        Result result = new Result();
        var prog = TestInstrBuilder.CreateProgram();
        bool res = parser.Process(result, scriptTokens, prog);

        //==> Check the result
        Assert.IsTrue(res);
        Assert.AreEqual(1, prog.ListInstr.Count);

        // OnExcel
        Assert.AreEqual(InstrType.OnExcel, prog.ListInstr[0].InstrType);
        InstrOnExcel instrOnExcel = prog.ListInstr[0] as InstrOnExcel;

        // InstrOnSheet
        InstrOnSheet instrOnSheet = instrOnExcel.ListSheets[0];

        // check IfThen
        InstrIfThenElse instrIfThenElse = instrOnSheet.ListInstrForEachRow[0] as InstrIfThenElse;

        // check If  -> bool expression
        InstrBoolExpr instrBoolExpr = instrIfThenElse.InstrIf.InstrBase as InstrBoolExpr;
        Assert.IsNotNull(instrBoolExpr);
        Assert.AreEqual(2, instrBoolExpr.ListOperand.Count);
        Assert.AreEqual(InstrBoolExprOperator.And, instrBoolExpr.Operator);

        // Comparison: A.Cell > 10
        InstrComparison instrComparison = instrBoolExpr.ListOperand[0] as InstrComparison;
        Assert.IsNotNull(instrComparison);
        InstrColCellFunc instrColCellFunc = instrComparison.OperandLeft as InstrColCellFunc;
        Assert.IsNotNull(instrColCellFunc);
        Assert.AreEqual("A", instrColCellFunc.ColName);
        // >
        Assert.AreEqual(SepComparisonOperator.GreaterThan, instrComparison.Operator.Operator);
    }

}
