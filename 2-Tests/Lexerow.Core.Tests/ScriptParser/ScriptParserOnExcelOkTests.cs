using FakeItEasy;
using Lexerow.Core.ScriptCompile.SyntaxAnalyze;
using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;
using Lexerow.Core.System.Compilator;
using Lexerow.Core.Tests._05_Common;

namespace Lexerow.Core.Tests.ScriptParser;

/// <summary>
/// Test script lexical analyzer on OnExcel instr.
/// Neither OnSheet, nor FirstRow.
/// One IfThen in ForEachRow.
/// </summary>
[TestClass]
public class ScriptParserOnExcelOkTests
{
    /// <summary>
    /// Compile: OnExcel, very short version
    /// Result: one instruction OnExcel
    /// Implicite: sheet=0, FirstRow=1
    /// 
    ///	OnExcel "file.xlsx"
    ///   ForEach Row
    ///     If A.Cell >10 Then A.Cell=10
    ///   Next 
    /// End OnExcel  
    /// </summary>
    [TestMethod]
    public void VeryShortOnExcelFileStringOk()
    {
        ScriptLineTokensTest lineTok;
        List<ScriptLineTokens> script = new List<ScriptLineTokens>();

        //-build one line of tokens
        lineTok = ScriptLineTokensTest.CreateOnExcelFileString("\"data.xlsx\"");
        script.Add(lineTok);

        // ForEach Row
        lineTok = new ScriptLineTokensTest();
        lineTok.AddTokenName(2, "ForEach", "Row");
        script.Add(lineTok);

        // If A.Cell >10 Then A.Cell=10
        TestScriptBuilder.BuidIfACellEq10ThenSetACell(3,script);

        // Next
        lineTok = new ScriptLineTokensTest();
        lineTok.AddTokenName(1, 1, "Next");
        script.Add(lineTok);

        // End OnExcel
        lineTok = new ScriptLineTokensTest();
        lineTok.AddTokenName(1, "End", "OnExcel");
        script.Add(lineTok);

        var logger = A.Fake<IActivityLogger>();
        Parser sa = new Parser(logger);

        //--parse the tokens
        ExecResult execResult = new ExecResult();
        bool res = sa.Process(execResult, script, out List<InstrBase> listInstr);

        Assert.IsTrue(res);
        Assert.AreEqual(1, listInstr.Count);

        //--check result
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

        // check If-Operator
        InstrSepComparison instrSepComparison = instrComparison.Operator;
        Assert.IsNotNull(instrSepComparison);
        Assert.AreEqual(SepComparisonOperator.GreaterThan, instrSepComparison.Operator);

        // check If-Operand Left
        TestBuilder.TestInstrColCellFuncValue("If-OperandLeft", instrComparison.OperandLeft, "A", 1);

        // check If-Operand Right
        TestBuilder.TestInstrConstValue("If-OperandRight", instrComparison.OperandRight, 10);

        // check Then, SetVar -> Left:InstrColCellFunc, Right InstrConstValue: 10
        Assert.IsNotNull(instrIfThenElse.InstrThen);
        Assert.AreEqual(1, instrIfThenElse.InstrThen.ListInstr.Count);

        InstrSetVar instrSetVar = instrIfThenElse.InstrThen.ListInstr[0] as InstrSetVar;
        Assert.IsNotNull(instrSetVar);

        TestBuilder.TestInstrColCellFuncValue("Then-SetVar-OperandLeft", instrSetVar.InstrLeft, "A", 1);
        TestBuilder.TestInstrConstValue("Then-SetVar-OperandRIght", instrSetVar.InstrRight, 10);

    }

    /// <summary>
    /// Compile:, OnExcel, very short version
    /// Result: one instruction OnExcel
    /// Implicite: sheet=0, FirstRow=1
    /// 
    /// file=OpenExcel("file.xlsx")
    ///	OnExcel file
    ///   ForEach Row
    ///     If A.Cell >10 Then A.Cell=10
    ///   Next 
    /// End OnExcel  
    /// </summary>
    [TestMethod]
    public void VeryShortOnExcelFileNameOk()
    {
        ScriptLineTokensTest lineTok;
        List<ScriptLineTokens> script = new List<ScriptLineTokens>();

        //--file=OpenExcel("data.xlsx")
        var line = TestScriptBuilder.BuildSelectFiles("file", "\"data.xlsx\"");
        script.Add(line);

        //-build one line of tokens
        lineTok = ScriptLineTokensTest.CreateOnExcelFileName("file");
        script.Add(lineTok);

        // ForEach Row
        lineTok = new ScriptLineTokensTest();
        lineTok.AddTokenName(2, "ForEach", "Row");
        script.Add(lineTok);

        // If A.Cell >10 Then A.Cell=10
        TestScriptBuilder.BuidIfACellEq10ThenSetACell(3, script);

        // Next
        lineTok = new ScriptLineTokensTest();
        lineTok.AddTokenName(1, 1, "Next");
        script.Add(lineTok);

        // End OnExcel
        lineTok = new ScriptLineTokensTest();
        lineTok.AddTokenName(1, "End", "OnExcel");
        script.Add(lineTok);

        var logger = A.Fake<IActivityLogger>();
        Parser sa = new Parser(logger);

        //--parse the tokens
        ExecResult execResult = new ExecResult();
        bool res = sa.Process(execResult, script, out List<InstrBase> listInstr);

        Assert.IsTrue(res);
        Assert.AreEqual(2, listInstr.Count);

        //--SetVar: file=OpenExcel("data.xlsx")
        Assert.AreEqual(InstrType.SetVar, listInstr[0].InstrType);
        InstrSetVar instrSetVar= listInstr[0] as InstrSetVar;
        // left:  InstrObjectName -> file
        InstrObjectName instrObjectName = instrSetVar.InstrLeft as InstrObjectName;
        Assert.IsNotNull(instrObjectName);
        Assert.AreEqual("file", instrObjectName.ObjectName);
        // right:  InstrOpenExcel
        InstrSelectFiles instrOpenExcel= instrSetVar.InstrRight as InstrSelectFiles;
        Assert.IsNotNull(instrOpenExcel);
        // no need to test further

        //--OnExcel
        Assert.AreEqual(InstrType.OnExcel, listInstr[1].InstrType);
        InstrOnExcel instrOnExcel = listInstr[1] as InstrOnExcel;

        // OnExcel.ListFiles: OnExcel file -> InstrObjectName
        Assert.IsNotNull(instrOnExcel.InstrFiles);
        instrObjectName = instrOnExcel.InstrFiles as InstrObjectName;
        Assert.AreEqual("file", instrObjectName.ObjectName);

        // no need to test further
    }



    /// <summary>
    /// Compile:, OnExcel, very short version
    /// Result: one instruction OnExcel
    /// Implicite: sheet=0, FirstRow=1
    /// 
    ///	OnExcel 
    ///	  "file.xlsx"
    ///        ForEach Row
    ///            If A.Cell >10 Then A.Cell=10
    ///         Next
    /// End OnExcel  
    /// </summary>
    [TestMethod]
    public void VeryShortOnExcelOneIfThenInlineOk()
    {
        List<ScriptLineTokens> script = new List<ScriptLineTokens>();
        ScriptLineTokensTest lineTok;

        //-OnExcel
        lineTok = new ScriptLineTokensTest();
        lineTok.AddTokenName(1, 1, "OnExcel");
        script.Add(lineTok);

        //-"data.xslx"
        lineTok = new ScriptLineTokensTest();
        lineTok.AddTokenString(1, 1, "\"data.xlsx\"");
        script.Add(lineTok);

        //-ForEach Row
        lineTok = new ScriptLineTokensTest();
        lineTok.AddTokenName(5, "ForEach", "Row");
        script.Add(lineTok);

        // If A.Cell >10 Then A.Cell=10
        TestScriptBuilder.BuidIfACellEq10ThenSetACell(3, script);

        // Next
        lineTok = new ScriptLineTokensTest();
        lineTok.AddTokenName(1, 1, "Next");
        script.Add(lineTok);

        // End OnExcel
        lineTok = new ScriptLineTokensTest();
        lineTok.AddTokenName(1, 1, "End");
        lineTok.AddTokenName(1, 1, "OnExcel");
        script.Add(lineTok);

        var logger = A.Fake<IActivityLogger>();
        Parser sa = new Parser(logger);

        ExecResult execResult = new ExecResult();
        bool res = sa.Process(execResult, script, out List<InstrBase> listInstr);

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

        // check If-Operator
        InstrSepComparison instrSepComparison = instrComparison.Operator as InstrSepComparison;
        Assert.IsNotNull(instrSepComparison);
        Assert.AreEqual(SepComparisonOperator.GreaterThan, instrSepComparison.Operator);

        // check If-Operand Left
        TestBuilder.TestInstrColCellFuncValue("If-OperandLeft", instrComparison.OperandLeft, "A", 1);

        // check If-Operand Right
        TestBuilder.TestInstrConstValue("If-OperandRight", instrComparison.OperandRight, 10);

        // check Then, SetVar -> Left:InstrColCellFunc, Right InstrConstValue: 10
        Assert.IsNotNull(instrIfThenElse.InstrThen);
        Assert.AreEqual(1, instrIfThenElse.InstrThen.ListInstr.Count);

        InstrSetVar instrSetVar = instrIfThenElse.InstrThen.ListInstr[0] as InstrSetVar;
        Assert.IsNotNull(instrSetVar);

        TestBuilder.TestInstrColCellFuncValue("Then-SetVar-OperandLeft", instrSetVar.InstrLeft, "A", 1);
        TestBuilder.TestInstrConstValue("Then-SetVar-OperandRIght", instrSetVar.InstrRight, 10);
    }

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
        List<ScriptLineTokens> script = new List<ScriptLineTokens>();
        ScriptLineTokensTest lineTok;

        //-OnExcel
        lineTok = new ScriptLineTokensTest();
        lineTok.AddTokenName(1, 1, "OnExcel");
        script.Add(lineTok);

        //-"data.xslx"
        lineTok = new ScriptLineTokensTest();
        lineTok.AddTokenString(2, 1, "\"data.xlsx\"");
        script.Add(lineTok);

        //-ForEach Row
        lineTok = new ScriptLineTokensTest();
        lineTok.AddTokenName(3, "ForEach", "Row");
        script.Add(lineTok);

        // If A.Cell >10 Then A.Cell=10
        TestScriptBuilder.BuidIfACellEq10ThenSetACell(3, script);

        // If B.Cell >12 Then B.Cell=12
        TestScriptBuilder.BuidIfACellEq10ThenSetACell(4, script,"B",">",12,"B",12);

        // Next
        lineTok = new ScriptLineTokensTest();
        lineTok.AddTokenName(1, 1, "Next");
        script.Add(lineTok);

        // End OnExcel
        lineTok = new ScriptLineTokensTest();
        lineTok.AddTokenName(1, 1, "End");
        lineTok.AddTokenName(1, 1, "OnExcel");
        script.Add(lineTok);

        var logger = A.Fake<IActivityLogger>();
        Parser sa = new Parser(logger);

        ExecResult execResult = new ExecResult();
        bool res = sa.Process(execResult, script, out List<InstrBase> listInstr);

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
        TestBuilder.TestInstrColCellFuncValue("If-OperandLeft", instrComparison.OperandLeft, "A", 1);

        // check If-Operand Right
        TestBuilder.TestInstrConstValue("If-OperandRight", instrComparison.OperandRight, 10);

        // check Then, SetVar -> Left:InstrColCellFunc, Right InstrConstValue: 10
        Assert.IsNotNull(instrIfThenElse.InstrThen);
        Assert.AreEqual(1, instrIfThenElse.InstrThen.ListInstr.Count);

        InstrSetVar instrSetVar = instrIfThenElse.InstrThen.ListInstr[0] as InstrSetVar;
        Assert.IsNotNull(instrSetVar);

        TestBuilder.TestInstrColCellFuncValue("Then-SetVar-OperandLeft", instrSetVar.InstrLeft, "A", 1);
        TestBuilder.TestInstrConstValue("Then-SetVar-OperandRIght", instrSetVar.InstrRight, 10);


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
        TestBuilder.TestInstrColCellFuncValue("If2-OperandLeft", instrComparison.OperandLeft, "B", 2);

        // check If-Operand Right
        TestBuilder.TestInstrConstValue("If2-OperandRight", instrComparison.OperandRight, 12);

        // check Then, SetVar -> Left:InstrColCellFunc, Right InstrConstValue: 10
        Assert.IsNotNull(instrIfThenElse.InstrThen);
        Assert.AreEqual(1, instrIfThenElse.InstrThen.ListInstr.Count);

        instrSetVar = instrIfThenElse.InstrThen.ListInstr[0] as InstrSetVar;
        Assert.IsNotNull(instrSetVar);

        TestBuilder.TestInstrColCellFuncValue("Then2-SetVar-OperandLeft", instrSetVar.InstrLeft, "B", 2);
        TestBuilder.TestInstrConstValue("Then-SetVar-OperandRIght", instrSetVar.InstrRight, 12);
    }

    /// <summary>
    /// Special case.
    /// check all token in one line:
    /// OnExcel "data.xlsx" ForEach Row If A.Cell >10 Then A.Cell=10 Next End OnExcel
    /// 
    /// </summary>
    [TestMethod]
    public void VeryShortOnExcelOneLineOk()
    {
        List<ScriptLineTokens> script = new List<ScriptLineTokens>();
        ScriptLineTokens lineTok;

        //-OnExcel
        lineTok = new ScriptLineTokens();
        lineTok.AddTokenName(1, 1, "OnExcel");
        lineTok.AddTokenString(1, 6, "\"data.xlsx\"");
        lineTok.AddTokenName(1, 12, "ForEach");
        lineTok.AddTokenName(1, 18, "Row");

        // If A.Cell >10 Then A.Cell=10
        TestScriptBuilder.BuidIfACellEq10ThenSetACell(1,lineTok,"A",">",10,"A", 10);
        script.Add(lineTok);
        lineTok.AddTokenName(1, 25, "Next");
        lineTok.AddTokenName(1, 34, "End");
        lineTok.AddTokenName(1, 40, "OnExcel");

        var logger = A.Fake<IActivityLogger>();
        Parser sa = new Parser(logger);

        ExecResult execResult = new ExecResult();
        bool res = sa.Process(execResult, script, out List<InstrBase> listInstr);

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

        // check If-Operator
        InstrSepComparison instrSepComparison = instrComparison.Operator as InstrSepComparison;
        Assert.IsNotNull(instrSepComparison);
        Assert.AreEqual(SepComparisonOperator.GreaterThan, instrSepComparison.Operator);

        // check If-Operand Left
        TestBuilder.TestInstrColCellFuncValue("If-OperandLeft", instrComparison.OperandLeft, "A", 1);

        // check If-Operand Right
        TestBuilder.TestInstrConstValue("If-OperandRight", instrComparison.OperandRight, 10);

        // check Then, SetVar -> Left:InstrColCellFunc, Right InstrConstValue: 10
        Assert.IsNotNull(instrIfThenElse.InstrThen);
        Assert.AreEqual(1, instrIfThenElse.InstrThen.ListInstr.Count);

        InstrSetVar instrSetVar = instrIfThenElse.InstrThen.ListInstr[0] as InstrSetVar;
        Assert.IsNotNull(instrSetVar);

        TestBuilder.TestInstrColCellFuncValue("Then-SetVar-OperandLeft", instrSetVar.InstrLeft, "A", 1);
        TestBuilder.TestInstrConstValue("Then-SetVar-OperandRIght", instrSetVar.InstrRight, 10);
    }

    /// <summary>
    /// Compile:, OnExcel, very short version
    /// Result: one instruction OnExcel
    /// Implicite: sheet=0, FirstRow=1
    /// 
    ///	OnExcel "file.xlsx" 
    ///        ForEach Row
    ///            If A.Cell >10 Then 
    ///               A.Cell=10
    ///            End If   
    ///         Next
    /// End OnExcel  
    /// </summary>
    [TestMethod]
    public void VeryShortOnExcelIfThenEndIfOk()
    {
        List<ScriptLineTokens> script = new List<ScriptLineTokens>();
        ScriptLineTokensTest lineTok;

        //-OnExcel "data.xslx"
        lineTok = ScriptLineTokensTest.CreateOnExcelFileString("\"data.xlsx\"");
        script.Add(lineTok);

        //-ForEach Row
        lineTok = new ScriptLineTokensTest();
        lineTok.AddTokenName(2, "ForEach", "Row");
        script.Add(lineTok);

        // If A.Cell >10 Then
        lineTok = new ScriptLineTokensTest();
        lineTok.AddTokenName(3, 1, "If");
        TestScriptBuilder.BuidColCellCompValue(3, lineTok, "A", ">", 10 );
        lineTok.AddTokenName(3, 1, "Then");
        script.Add(lineTok);

        // A.Cell=10
        lineTok = new ScriptLineTokensTest();
        TestScriptBuilder.BuidColCellCompValue(4, lineTok,"A","=", 10);
        script.Add(lineTok);

        // End If
        lineTok = new ScriptLineTokensTest();
        lineTok.AddTokenName(5, "End", "If");
        script.Add(lineTok);

        // Next
        lineTok = new ScriptLineTokensTest();
        lineTok.AddTokenName(6, 1, "Next");
        script.Add(lineTok);

        // End OnExcel
        lineTok = new ScriptLineTokensTest();
        lineTok.AddTokenName(7, "End", "OnExcel");
        script.Add(lineTok);

        var logger = A.Fake<IActivityLogger>();
        Parser sa = new Parser(logger);

        ExecResult execResult = new ExecResult();
        bool res = sa.Process(execResult, script, out List<InstrBase> listInstr);

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

        // check If-Operator
        InstrSepComparison instrSepComparison = instrComparison.Operator as InstrSepComparison;
        Assert.IsNotNull(instrSepComparison);
        Assert.AreEqual(SepComparisonOperator.GreaterThan, instrSepComparison.Operator);

        // check If-Operand Left
        TestBuilder.TestInstrColCellFuncValue("If-OperandLeft", instrComparison.OperandLeft, "A", 1);

        // check If-Operand Right
        TestBuilder.TestInstrConstValue("If-OperandRight", instrComparison.OperandRight, 10);

        // check Then, SetVar -> Left:InstrColCellFunc, Right InstrConstValue: 10
        Assert.IsNotNull(instrIfThenElse.InstrThen);
        Assert.AreEqual(1, instrIfThenElse.InstrThen.ListInstr.Count);

        InstrSetVar instrSetVar = instrIfThenElse.InstrThen.ListInstr[0] as InstrSetVar;
        Assert.IsNotNull(instrSetVar);

        TestBuilder.TestInstrColCellFuncValue("Then-SetVar-OperandLeft", instrSetVar.InstrLeft, "A", 1);
        TestBuilder.TestInstrConstValue("Then-SetVar-OperandRIght", instrSetVar.InstrRight, 10);
    }

}
