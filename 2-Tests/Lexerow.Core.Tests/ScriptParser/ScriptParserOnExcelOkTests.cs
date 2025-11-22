using FakeItEasy;
using Lexerow.Core.ScriptCompile.Parse;
using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;
using Lexerow.Core.System.ScriptDef;
using Lexerow.Core.Tests._05_Common;

namespace Lexerow.Core.Tests.ScriptParser;

/// <summary>
/// Test script parser on OnExcel instr.
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
        int numLine=0;
        List<ScriptLineTokens> scriptTokens = new List<ScriptLineTokens>();

        // OnExcel "data.xlsx"
        TestTokensBuilder.AddLineOnExcelFileString(numLine++, scriptTokens, "\"data.xlsx\"");

        // ForEach Row
        TestTokensBuilder.AddLineForEachRow(numLine++, scriptTokens);

        // If A.Cell >10 Then A.Cell=10
        TestTokensBuilder.BuidIfColCellEqualIntThenSetColCellInt(numLine++, scriptTokens);

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
        InstrSepComparison instrSepComparison = instrComparison.Operator;
        Assert.IsNotNull(instrSepComparison);
        Assert.AreEqual(SepComparisonOperator.GreaterThan, instrSepComparison.Operator);

        // check If-Operand Left
        TestInstrHelper.TestInstrColCellFuncValue("If-OperandLeft", instrComparison.OperandLeft, "A", 1);

        // check If-Operand Right
        TestInstrHelper.TestInstrValue("If-OperandRight", instrComparison.OperandRight, 10);

        // check Then, SetVar -> Left:InstrColCellFunc, Right InstrValue: 10
        Assert.IsNotNull(instrIfThenElse.InstrThen);
        Assert.AreEqual(1, instrIfThenElse.InstrThen.ListInstr.Count);

        InstrSetVar instrSetVar = instrIfThenElse.InstrThen.ListInstr[0] as InstrSetVar;
        Assert.IsNotNull(instrSetVar);

        TestInstrHelper.TestInstrColCellFuncValue("Then-SetVar-OperandLeft", instrSetVar.InstrLeft, "A", 1);
        TestInstrHelper.TestInstrValue("Then-SetVar-OperandRIght", instrSetVar.InstrRight, 10);
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
        int numLine = 0;
        List<ScriptLineTokens> scriptTokens = new List<ScriptLineTokens>();

        //--file=OpenExcel("data.xlsx")
        TestTokensBuilder.AddLineSelectFiles(numLine++, scriptTokens, "file", "\"data.xlsx\"");

        // OnExcel file
        TestTokensBuilder.CreateOnExcelFileName(numLine++, scriptTokens, "file");

        // ForEach Row
        TestTokensBuilder.AddLineForEachRow(numLine++, scriptTokens); 

        // If A.Cell >10 Then A.Cell=10
        TestTokensBuilder.BuidIfColCellEqualIntThenSetColCellInt(3, scriptTokens);

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
        Assert.AreEqual(2, prog.ListInstr.Count);

        //--SetVar: file=OpenExcel("data.xlsx")
        Assert.AreEqual(InstrType.SetVar, prog.ListInstr[0].InstrType);
        InstrSetVar instrSetVar = prog.ListInstr[0] as InstrSetVar;
        // left:  InstrObjectName -> file
        InstrObjectName instrObjectName = instrSetVar.InstrLeft as InstrObjectName;
        Assert.IsNotNull(instrObjectName);
        Assert.AreEqual("file", instrObjectName.ObjectName);
        // right:  InstrOpenExcel
        InstrSelectFiles instrOpenExcel = instrSetVar.InstrRight as InstrSelectFiles;
        Assert.IsNotNull(instrOpenExcel);
        // no need to test further

        //--OnExcel
        Assert.AreEqual(InstrType.OnExcel, prog.ListInstr[1].InstrType);
        InstrOnExcel instrOnExcel = prog.ListInstr[1] as InstrOnExcel;

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
        List<ScriptLineTokens> scriptTokens = new List<ScriptLineTokens>();
        ScriptLineTokens line;

        //-OnExcel
        line = new ScriptLineTokens();
        line.AddTokenName(1, 1, "OnExcel");
        scriptTokens.Add(line);

        //-"data.xslx"
        line = new ScriptLineTokens();
        line.AddTokenString(1, 1, "\"data.xlsx\"");
        scriptTokens.Add(line);

        //-ForEach Row
        line = new ScriptLineTokens();
        TestTokensBuilder.AddTokenName(5, line, "ForEach", "Row");
        scriptTokens.Add(line);

        // If A.Cell >10 Then A.Cell=10
        TestTokensBuilder.BuidIfColCellEqualIntThenSetColCellInt(3, scriptTokens);

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
        TestInstrHelper.TestInstrColCellFuncValue("If-OperandLeft", instrComparison.OperandLeft, "A", 1);

        // check If-Operand Right
        TestInstrHelper.TestInstrValue("If-OperandRight", instrComparison.OperandRight, 10);

        // check Then, SetVar -> Left:InstrColCellFunc, Right InstrValue: 10
        Assert.IsNotNull(instrIfThenElse.InstrThen);
        Assert.AreEqual(1, instrIfThenElse.InstrThen.ListInstr.Count);

        InstrSetVar instrSetVar = instrIfThenElse.InstrThen.ListInstr[0] as InstrSetVar;
        Assert.IsNotNull(instrSetVar);

        TestInstrHelper.TestInstrColCellFuncValue("Then-SetVar-OperandLeft", instrSetVar.InstrLeft, "A", 1);
        TestInstrHelper.TestInstrValue("Then-SetVar-OperandRIght", instrSetVar.InstrRight, 10);
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
        TestInstrHelper.TestInstrColCellFuncValue("If-OperandLeft", instrComparison.OperandLeft, "A", 1);

        // check If-Operand Right
        TestInstrHelper.TestInstrValue("If-OperandRight", instrComparison.OperandRight, 10);

        // check Then, SetVar -> Left:InstrColCellFunc, Right InstrValue: 10
        Assert.IsNotNull(instrIfThenElse.InstrThen);
        Assert.AreEqual(1, instrIfThenElse.InstrThen.ListInstr.Count);

        InstrSetVar instrSetVar = instrIfThenElse.InstrThen.ListInstr[0] as InstrSetVar;
        Assert.IsNotNull(instrSetVar);

        TestInstrHelper.TestInstrColCellFuncValue("Then-SetVar-OperandLeft", instrSetVar.InstrLeft, "A", 1);
        TestInstrHelper.TestInstrValue("Then-SetVar-OperandRIght", instrSetVar.InstrRight, 10);

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
        TestInstrHelper.TestInstrColCellFuncValue("If2-OperandLeft", instrComparison.OperandLeft, "B", 2);

        // check If-Operand Right
        TestInstrHelper.TestInstrValue("If2-OperandRight", instrComparison.OperandRight, 12);

        // check Then, SetVar -> Left:InstrColCellFunc, Right InstrValue: 10
        Assert.IsNotNull(instrIfThenElse.InstrThen);
        Assert.AreEqual(1, instrIfThenElse.InstrThen.ListInstr.Count);

        instrSetVar = instrIfThenElse.InstrThen.ListInstr[0] as InstrSetVar;
        Assert.IsNotNull(instrSetVar);

        TestInstrHelper.TestInstrColCellFuncValue("Then2-SetVar-OperandLeft", instrSetVar.InstrLeft, "B", 2);
        TestInstrHelper.TestInstrValue("Then-SetVar-OperandRIght", instrSetVar.InstrRight, 12);
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
        List<ScriptLineTokens> scriptTokens = new List<ScriptLineTokens>();
        ScriptLineTokens lineTok;

        //-OnExcel
        lineTok = new ScriptLineTokens();
        lineTok.AddTokenName(1, 1, "OnExcel");
        lineTok.AddTokenString(1, 6, "\"data.xlsx\"");
        lineTok.AddTokenName(1, 12, "ForEach");
        lineTok.AddTokenName(1, 18, "Row");

        // If A.Cell >10 Then A.Cell=10
        TestTokensBuilder.BuidIfColCellCompIntThenSetColCellInt(1, lineTok, "A", ">", 10, "A", 10);
        scriptTokens.Add(lineTok);
        lineTok.AddTokenName(1, 25, "Next");
        lineTok.AddTokenName(1, 34, "End");
        lineTok.AddTokenName(1, 40, "OnExcel");

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
        TestInstrHelper.TestInstrColCellFuncValue("If-OperandLeft", instrComparison.OperandLeft, "A", 1);

        // check If-Operand Right
        TestInstrHelper.TestInstrValue("If-OperandRight", instrComparison.OperandRight, 10);

        // check Then, SetVar -> Left:InstrColCellFunc, Right InstrValue: 10
        Assert.IsNotNull(instrIfThenElse.InstrThen);
        Assert.AreEqual(1, instrIfThenElse.InstrThen.ListInstr.Count);

        InstrSetVar instrSetVar = instrIfThenElse.InstrThen.ListInstr[0] as InstrSetVar;
        Assert.IsNotNull(instrSetVar);

        TestInstrHelper.TestInstrColCellFuncValue("Then-SetVar-OperandLeft", instrSetVar.InstrLeft, "A", 1);
        TestInstrHelper.TestInstrValue("Then-SetVar-OperandRIght", instrSetVar.InstrRight, 10);
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
    public void VeryShortOnExcelIfThenEndIfOk()
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
        TestInstrHelper.TestInstrColCellFuncValue("If-OperandLeft", instrComparison.OperandLeft, "A", 1);

        // check If-Operand Right
        TestInstrHelper.TestInstrValue("If-OperandRight", instrComparison.OperandRight, 10);

        // check Then, SetVar -> Left:InstrColCellFunc, Right InstrValue: 10
        Assert.IsNotNull(instrIfThenElse.InstrThen);
        Assert.AreEqual(1, instrIfThenElse.InstrThen.ListInstr.Count);

        InstrSetVar instrSetVar = instrIfThenElse.InstrThen.ListInstr[0] as InstrSetVar;
        Assert.IsNotNull(instrSetVar);

        TestInstrHelper.TestInstrColCellFuncValue("Then-SetVar-OperandLeft", instrSetVar.InstrLeft, "A", 1);
        TestInstrHelper.TestInstrValue("Then-SetVar-OperandRIght", instrSetVar.InstrRight, 10);
    }
}