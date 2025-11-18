using FakeItEasy;
using Lexerow.Core.ScriptCompile.Parse;
using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;
using Lexerow.Core.System.ScriptDef;
using Lexerow.Core.Tests._05_Common;

namespace Lexerow.Core.Tests.ScriptParser;

/// <summary>
/// Test script parser on OnExcel instr.
/// Focus on If A.Cell=blank and A.Cell=null.
/// One IfThen in ForEachRow.
/// </summary>
[TestClass]
public class ScriptParserOnExcelBlankNullTests
{
    /// <summary>
    /// Implicite: sheet=0, FirstRow=1
    ///
    ///	OnExcel "file.xlsx"
    ///   ForEach Row
    ///     If A.Cell=blank Then A.Cell=10
    ///   Next
    /// End OnExcel
    /// </summary>
    [TestMethod]
    public void OnExcelIfACellEqualBlankOk()
    {
        int numLine = 1;
        List<ScriptLineTokens> script = new List<ScriptLineTokens>();

        // OnExcel "data.xlsx"
        TestTokensBuilder.AddLineOnExcelFileString(numLine++, script, "\"data.xlsx\"");

        // ForEach Row
        TestTokensBuilder.AddLineForEachRow(numLine++, script);

        // If A.Cell=blank Then A.Cell=12
        TestTokensBuilder.BuidIfColCellCompKeywordThenSetColCellInt(3, script, "A", "=", "Blank", "A", 12);

        // Next
        TestTokensBuilder.AddLineNext(numLine++, script);

        // End OnExcel
        TestTokensBuilder.AddLineEndOnExcel(numLine++, script);

        //==>just to check the content of the script
        //var scriptCheck = TestTokens2ScriptBuilder.BuildScript(script);

        //==> Parse the script tokens
        var logger = A.Fake<IActivityLogger>();
        Parser sa = new Parser(logger);
        ExecResult execResult = new ExecResult();
        bool res = sa.Process(execResult, script, out List<InstrBase> listInstr);

        //==> Check result
        Assert.IsTrue(res);
        Assert.AreEqual(1, listInstr.Count);

        // OnExcel
        Assert.AreEqual(InstrType.OnExcel, listInstr[0].InstrType);
        InstrOnExcel instrOnExcel = listInstr[0] as InstrOnExcel;

        // OnExcel.ListFiles
        Assert.IsNotNull(instrOnExcel.InstrFiles);
        InstrConstValue constExcelFileName = instrOnExcel.InstrFiles as InstrConstValue;
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

        // check If-Operator =
        InstrSepComparison instrSepComparison = instrComparison.Operator;
        Assert.IsNotNull(instrSepComparison);
        Assert.AreEqual(SepComparisonOperator.Equal, instrSepComparison.Operator);

        // check If-Operand Left -> If A.Cell
        InstrTestHelper.TestInstrColCellFuncValue("If-OperandLeft", instrComparison.OperandLeft, "A", 1);

        // check If-Operand Right  -> Blank
        InstrBlank instrBlank = instrComparison.OperandRight as InstrBlank;
        Assert.IsNotNull(instrBlank);
        //TestBuilder.TestInstrKeyword("If-OperandRight", instrComparison.OperandRight, 10);

        // check Then, SetVar -> Left:InstrColCellFunc, Right InstrConstValue: 12
        Assert.IsNotNull(instrIfThenElse.InstrThen);
        Assert.AreEqual(1, instrIfThenElse.InstrThen.ListInstr.Count);

        InstrSetVar instrSetVar = instrIfThenElse.InstrThen.ListInstr[0] as InstrSetVar;
        Assert.IsNotNull(instrSetVar);

        InstrTestHelper.TestInstrColCellFuncValue("Then-SetVar-OperandLeft", instrSetVar.InstrLeft, "A", 1);
        InstrTestHelper.TestInstrConstValue("Then-SetVar-OperandRIght", instrSetVar.InstrRight, 12);
    }

    /// <summary>
    /// Implicite: sheet=0, FirstRow=1
    ///
    ///	OnExcel "file.xlsx"
    ///   ForEach Row
    ///     If A.Cell>blank Then A.Cell=10
    ///   Next
    /// End OnExcel
    /// </summary>
    [TestMethod]
    public void OnExcelIfACellGreaterBlankError()
    {
        int numLine = 0;
        List<ScriptLineTokens> script = new List<ScriptLineTokens>();

        // OnExcel "data.xlsx"
        TestTokensBuilder.AddLineOnExcelFileString(numLine++, script, "\"data.xlsx\"");

        // ForEach Row
        TestTokensBuilder.AddLineForEachRow(numLine++, script);  

        // If A.Cell=blank Then A.Cell=12
        TestTokensBuilder.BuidIfColCellCompKeywordThenSetColCellInt(numLine++, script, "A", ">", "Blank", "A", 12);

        // Next
        TestTokensBuilder.AddLineNext(numLine++, script);

        // End OnExcel
        TestTokensBuilder.AddLineEndOnExcel(numLine++, script);

        //==>just to check the content of the script
        var scriptCheck = TestTokens2ScriptBuilder.BuildScript(script);

        //==> Parse the script tokens
        var logger = A.Fake<IActivityLogger>();
        Parser sa = new Parser(logger);

        ExecResult execResult = new ExecResult();
        bool res = sa.Process(execResult, script, out List<InstrBase> listInstr);

        Assert.IsFalse(res);
        Assert.AreEqual(ErrorCode.ParserSepComparatorWrong, execResult.ListError[0].ErrorCode);
        Assert.AreEqual(">", execResult.ListError[0].Param);
        Assert.AreEqual(2, execResult.ListError[0].LineNum);
    }

    /// <summary>
    /// Default: sheet=0, FirstRow=1
    ///
    ///	OnExcel "file.xlsx"
    ///   ForEach Row
    ///     If A.Cell=9 Then A.Cell=Blank
    ///   Next
    /// End OnExcel
    /// </summary>
    [TestMethod]
    public void OnExcelThenACellEqBlankOk()
    {
        int numLine=0;
        List<ScriptLineTokens> script = new List<ScriptLineTokens>();

        // OnExcel "data.xlsx"
        TestTokensBuilder.AddLineOnExcelFileString(numLine++, script, "\"data.xlsx\"");

        // ForEach Row
        TestTokensBuilder.AddLineForEachRow(numLine++, script);

        // If A.Cell=9 Then A.Cell=Blank
        TestTokensBuilder.BuidIfColCellCompIntThenSetColCellKeyword(numLine++, script, "A", "=", 9, "A", "Blank");

        // Next
        TestTokensBuilder.AddLineNext(numLine++, script);

        // End OnExcel
        TestTokensBuilder.AddLineEndOnExcel(numLine++, script);

        //==>just to check the content of the script
        var scriptCheck = TestTokens2ScriptBuilder.BuildScript(script);

        //==> Parse the script tokens
        var logger = A.Fake<IActivityLogger>();
        Parser sa = new Parser(logger);

        ExecResult execResult = new ExecResult();
        bool res = sa.Process(execResult, script, out List<InstrBase> listInstr);

        Assert.IsTrue(res);
        Assert.AreEqual(1, listInstr.Count);

        //==> Check result
        // OnExcel
        Assert.AreEqual(InstrType.OnExcel, listInstr[0].InstrType);
        InstrOnExcel instrOnExcel = listInstr[0] as InstrOnExcel;

        // OnExcel.ListFiles
        Assert.IsNotNull(instrOnExcel.InstrFiles);
        InstrConstValue constExcelFileName = instrOnExcel.InstrFiles as InstrConstValue;
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

        // check If-Operator =
        InstrSepComparison instrSepComparison = instrComparison.Operator;
        Assert.IsNotNull(instrSepComparison);
        Assert.AreEqual(SepComparisonOperator.Equal, instrSepComparison.Operator);

        // check If-Operand Left -> If A.Cell
        InstrTestHelper.TestInstrColCellFuncValue("If-OperandLeft", instrComparison.OperandLeft, "A", 1);

        // check If-Operand Right
        InstrTestHelper.TestInstrConstValue("If-OperandRight", instrComparison.OperandRight, 9);

        // check Then, SetVar
        Assert.IsNotNull(instrIfThenElse.InstrThen);
        Assert.AreEqual(1, instrIfThenElse.InstrThen.ListInstr.Count);

        InstrSetVar instrSetVar = instrIfThenElse.InstrThen.ListInstr[0] as InstrSetVar;
        Assert.IsNotNull(instrSetVar);

        InstrTestHelper.TestInstrColCellFuncValue("Then-SetVar-OperandLeft", instrSetVar.InstrLeft, "A", 1);
        InstrBlank instrBlank = instrSetVar.InstrRight as InstrBlank;
        Assert.IsNotNull(instrBlank);
    }

    /// <summary>
    /// Default: sheet=0, FirstRow=1
    ///
    ///	OnExcel "file.xlsx"
    ///   ForEach Row
    ///     If A.Cell=9 Then A.Cell=Null
    ///   Next
    /// End OnExcel
    /// </summary>
    [TestMethod]
    public void OnExcelThenACellEqNullOk()
    {
        int numLine = 0;
        List<ScriptLineTokens> script = new List<ScriptLineTokens>();

        // OnExcel "data.xlsx"
        TestTokensBuilder.AddLineOnExcelFileString(numLine++, script, "\"data.xlsx\"");

        // ForEach Row
        TestTokensBuilder.AddLineForEachRow(numLine++, script);

        // If A.Cell=9 Then A.Cell=Blank
        TestTokensBuilder.BuidIfColCellCompIntThenSetColCellKeyword(3, script, "A", "=", 9, "A", "Null");

        // Next
        TestTokensBuilder.AddLineNext(numLine++, script);

        // End OnExcel
        TestTokensBuilder.AddLineEndOnExcel(numLine++, script); 

        //==>just to check the content of the script
        var scriptCheck = TestTokens2ScriptBuilder.BuildScript(script);

        //==> Parse the script tokens
        var logger = A.Fake<IActivityLogger>();
        Parser sa = new Parser(logger);

        ExecResult execResult = new ExecResult();
        bool res = sa.Process(execResult, script, out List<InstrBase> listInstr);

        Assert.IsTrue(res);
        Assert.AreEqual(1, listInstr.Count);

        //==> Check result
        // OnExcel
        Assert.AreEqual(InstrType.OnExcel, listInstr[0].InstrType);
        InstrOnExcel instrOnExcel = listInstr[0] as InstrOnExcel;

        // OnExcel.ListFiles
        Assert.IsNotNull(instrOnExcel.InstrFiles);
        InstrConstValue constExcelFileName = instrOnExcel.InstrFiles as InstrConstValue;
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

        // check If-Operator =
        InstrSepComparison instrSepComparison = instrComparison.Operator;
        Assert.IsNotNull(instrSepComparison);
        Assert.AreEqual(SepComparisonOperator.Equal, instrSepComparison.Operator);

        // check If-Operand Left -> If A.Cell
        InstrTestHelper.TestInstrColCellFuncValue("If-OperandLeft", instrComparison.OperandLeft, "A", 1);

        // check If-Operand Right
        InstrTestHelper.TestInstrConstValue("If-OperandRight", instrComparison.OperandRight, 9);

        // check Then, SetVar
        Assert.IsNotNull(instrIfThenElse.InstrThen);
        Assert.AreEqual(1, instrIfThenElse.InstrThen.ListInstr.Count);

        InstrSetVar instrSetVar = instrIfThenElse.InstrThen.ListInstr[0] as InstrSetVar;
        Assert.IsNotNull(instrSetVar);

        InstrTestHelper.TestInstrColCellFuncValue("Then-SetVar-OperandLeft", instrSetVar.InstrLeft, "A", 1);
        InstrNull instrNull = instrSetVar.InstrRight as InstrNull;
        Assert.IsNotNull(instrNull);
    }

    // a=12
    // If a=blank  -> error
}