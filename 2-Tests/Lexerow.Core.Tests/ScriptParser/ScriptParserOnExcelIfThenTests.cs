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
/// Focus on If-Then instr.
/// </summary>
[TestClass]
public class ScriptParserOnExcelIfThenTests
{
    /// <summary>
    /// Compile:, OnExcel, very short version
    /// Result: one instruction OnExcel
    /// Implicite: sheet=0, FirstRow=1
    ///
    ///	OnExcel
    ///	  "file.xlsx"
    ///        ForEach Row
    ///            If A.Cell >10 Then A.Cell=10
    ///            If B.Cell >12 Then B.Cell=12
    ///         Next
    /// End OnExcel
    /// </summary>
    [TestMethod]
    public void VeryShortOnExcel2IfThenInlineOk()
    {
        List<ScriptLineTokens> scriptTokens = new List<ScriptLineTokens>();
        ScriptLineTokens line;

        //-OnExcel
        line = new ScriptLineTokens();
        line.AddTokenName(1, 1, "OnExcel");
        scriptTokens.Add(line);

        //-"data.xslx"
        line = new ScriptLineTokens();
        line.AddTokenString(2, 1, "\"data.xlsx\"");
        scriptTokens.Add(line);

        //-ForEach Row
        line = new ScriptLineTokens();
        TestTokensBuilder.AddTokenName(3, line, "ForEach", "Row");
        scriptTokens.Add(line);

        // If A.Cell >10 Then A.Cell=10
        TestTokensBuilder.BuidIfColCellEqualIntThenSetColCellInt(3, scriptTokens);

        // If B.Cell >12 Then B.Cell=12
        TestTokensBuilder.BuidIfColCellCompIntThenSetColCellInt(4, scriptTokens, "B", ">", 12, "B", 12);

        // Next
        line = new ScriptLineTokens();
        line.AddTokenName(1, 1, "Next");
        scriptTokens.Add(line);

        // End OnExcel
        line = new ScriptLineTokens();
        line.AddTokenName(1, 1, "End");
        line.AddTokenName(1, 1, "OnExcel");
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
        Assert.AreEqual(1, prog.ListInstr.Count);

        // OnExcel
        Assert.AreEqual(InstrType.OnExcel, prog.ListInstr[0].InstrType);
        InstrOnExcel instrOnExcel = prog.ListInstr[0] as InstrOnExcel;

        // OnExcel.ListFiles
        Assert.IsNotNull(instrOnExcel.InstrFiles);
        InstrValue constExcelFileName = instrOnExcel.InstrFiles as InstrValue;
        Assert.AreEqual("data.xlsx", (constExcelFileName.ValueBase as ValueString).Val);

        // check InstrOnSheet
        Assert.AreEqual(1, instrOnExcel.ListSheets.Count);
        InstrOnSheet instrOnSheet = instrOnExcel.ListSheets[0];
        Assert.AreEqual(1, instrOnSheet.SheetNum);

        //--check IfThen -> 2 instr!
        Assert.AreEqual(2, instrOnSheet.ListInstrForEachRow.Count);

        //--check IfThen #1
        InstrIfThenElse instrIfThenElse = instrOnSheet.ListInstrForEachRow[0] as InstrIfThenElse;
        Assert.IsNotNull(instrIfThenElse);

        // check If
        Assert.IsNotNull(instrIfThenElse.InstrIf);
        InstrComparison instrComparison = instrIfThenElse.InstrIf.InstrBase as InstrComparison;
        Assert.IsNotNull(instrComparison);

        Assert.IsNotNull(instrComparison.OperandLeft);
        Assert.IsNotNull(instrComparison.OperandRight);
        Assert.IsNotNull(instrComparison.Operator);

        // check If-Operator
        InstrSepComparison instrSepComparison = instrComparison.Operator as InstrSepComparison;
        Assert.IsNotNull(instrSepComparison);
        Assert.AreEqual(SepComparisonOperator.GreaterThan, instrSepComparison.Operator);

        // check If-Operand Left
        Assert.IsTrue(TestInstrHelper.TestInstrColCellFuncValue(instrComparison.OperandLeft, "A", 1));

        // check If-Operand Right
        Assert.IsTrue(TestInstrHelper.TestInstrValue(instrComparison.OperandRight, 10));

        // check Then, SetVar -> Left:InstrColCellFunc, Right InstrValue: 10
        Assert.IsNotNull(instrIfThenElse.InstrThen);
        Assert.AreEqual(1, instrIfThenElse.InstrThen.ListInstr.Count);

        InstrSetVar instrSetVar = instrIfThenElse.InstrThen.ListInstr[0] as InstrSetVar;
        Assert.IsNotNull(instrSetVar);

        Assert.IsTrue(TestInstrHelper.TestInstrColCellFuncValue(instrSetVar.InstrLeft, "A", 1));
        Assert.IsTrue(TestInstrHelper.TestInstrValue(instrSetVar.InstrRight, 10));

        //--check IfThen #2=======
        instrIfThenElse = instrOnSheet.ListInstrForEachRow[1] as InstrIfThenElse;
        Assert.IsNotNull(instrIfThenElse);

        // check If
        Assert.IsNotNull(instrIfThenElse.InstrIf);
        instrComparison = instrIfThenElse.InstrIf.InstrBase as InstrComparison;
        Assert.IsNotNull(instrComparison);
        Assert.IsNotNull(instrComparison.OperandLeft);
        Assert.IsNotNull(instrComparison.OperandRight);
        Assert.IsNotNull(instrComparison.Operator);

        // check If-Operator
        instrSepComparison = instrComparison.Operator as InstrSepComparison;
        Assert.IsNotNull(instrSepComparison);
        Assert.AreEqual(SepComparisonOperator.GreaterThan, instrSepComparison.Operator);

        // check If-Operand Left
        Assert.IsTrue(TestInstrHelper.TestInstrColCellFuncValue(instrComparison.OperandLeft, "B", 2));

        // check If-Operand Right
        Assert.IsTrue(TestInstrHelper.TestInstrValue(instrComparison.OperandRight, 12));

        // check Then, SetVar -> Left:InstrColCellFunc, Right InstrValue: 10
        Assert.IsNotNull(instrIfThenElse.InstrThen);
        Assert.AreEqual(1, instrIfThenElse.InstrThen.ListInstr.Count);

        instrSetVar = instrIfThenElse.InstrThen.ListInstr[0] as InstrSetVar;
        Assert.IsNotNull(instrSetVar);

        Assert.IsTrue(TestInstrHelper.TestInstrColCellFuncValue(instrSetVar.InstrLeft, "B", 2));
        Assert.IsTrue(TestInstrHelper.TestInstrValue(instrSetVar.InstrRight, 12));
    }

    /// <summary>
    /// Compile:, OnExcel, very short version
    /// Result: one instruction OnExcel
    /// Implicite: sheet=0, FirstRow=1
    ///
    ///	OnExcel "data.xlsx"
    ///        ForEach Row
    ///            If A.Cell >10 Then
    ///               A.Cell=10
    ///            End If
    ///         Next
    /// End OnExcel
    /// </summary>
    [TestMethod]
    public void OnExcelIfThenEndIfOk()
    {
        int numLine = 0;
        List<ScriptLineTokens> scriptTokens = new List<ScriptLineTokens>();
        ScriptLineTokens line;

        // OnExcel "data.xlsx"
        TestTokensBuilder.AddLineOnExcelFileString(numLine++, scriptTokens, "\"data.xlsx\"");

        // ForEach Row
        TestTokensBuilder.AddLineForEachRow(numLine++, scriptTokens);

        // If A.Cell >10 Then
        line = new ScriptLineTokens();
        line.AddTokenName(3, 1, "If");
        TestTokensBuilder.BuidColCellOperInt(3, line, "A", ">", 10);
        line.AddTokenName(3, 1, "Then");
        scriptTokens.Add(line);

        // A.Cell=10
        line = new ScriptLineTokens();
        TestTokensBuilder.BuidColCellOperInt(4, line, "A", "=", 10);
        scriptTokens.Add(line);

        // End If
        line = new ScriptLineTokens();
        TestTokensBuilder.AddTokenName(5, line, "End", "If");
        scriptTokens.Add(line);

        // Next
        TestTokensBuilder.AddLineNext(numLine++, scriptTokens);

        // End OnExcel
        TestTokensBuilder.AddLineEndOnExcel(numLine++, scriptTokens);

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

        // OnExcel
        Assert.AreEqual(InstrType.OnExcel, prog.ListInstr[0].InstrType);
        InstrOnExcel instrOnExcel = prog.ListInstr[0] as InstrOnExcel;

        // OnExcel.ListFiles
        Assert.IsNotNull(instrOnExcel.InstrFiles);
        InstrValue constExcelFileName = instrOnExcel.InstrFiles as InstrValue;
        Assert.AreEqual("data.xlsx", (constExcelFileName.ValueBase as ValueString).Val);

        // check InstrOnSheet
        Assert.AreEqual(1, instrOnExcel.ListSheets.Count);
        InstrOnSheet instrOnSheet = instrOnExcel.ListSheets[0];
        Assert.AreEqual(1, instrOnSheet.SheetNum);

        // check IfThen
        Assert.AreEqual(1, instrOnSheet.ListInstrForEachRow.Count);
        InstrIfThenElse instrIfThenElse = instrOnSheet.ListInstrForEachRow[0] as InstrIfThenElse;
        Assert.IsNotNull(instrIfThenElse);

        // check If
        Assert.IsNotNull(instrIfThenElse.InstrIf);
        InstrComparison instrComparison = instrIfThenElse.InstrIf.InstrBase as InstrComparison;
        Assert.IsNotNull(instrComparison);
        Assert.IsNotNull(instrComparison.OperandLeft);
        Assert.IsNotNull(instrComparison.OperandRight);
        Assert.IsNotNull(instrComparison.Operator);

        // check If-Operator
        InstrSepComparison instrSepComparison = instrComparison.Operator as InstrSepComparison;
        Assert.IsNotNull(instrSepComparison);
        Assert.AreEqual(SepComparisonOperator.GreaterThan, instrSepComparison.Operator);

        // check If-Operand Left
        Assert.IsTrue(TestInstrHelper.TestInstrColCellFuncValue(instrComparison.OperandLeft, "A", 1));

        // check If-Operand Right
        Assert.IsTrue(TestInstrHelper.TestInstrValue(instrComparison.OperandRight, 10));

        // check Then, SetVar -> Left:InstrColCellFunc, Right InstrValue: 10
        Assert.IsNotNull(instrIfThenElse.InstrThen);
        Assert.AreEqual(1, instrIfThenElse.InstrThen.ListInstr.Count);

        InstrSetVar instrSetVar = instrIfThenElse.InstrThen.ListInstr[0] as InstrSetVar;
        Assert.IsNotNull(instrSetVar);

        Assert.IsTrue(TestInstrHelper.TestInstrColCellFuncValue(instrSetVar.InstrLeft, "A", 1));
        Assert.IsTrue(TestInstrHelper.TestInstrValue(instrSetVar.InstrRight, 10));
    }
}
