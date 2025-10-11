using Lexerow.Core.Scripts.SyntaxAnalyze;
using Lexerow.Core.System;
using Lexerow.Core.System.Compilator;
using Lexerow.Core.Tests._05_Common;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Tests.ScriptSyntaxAnalyze;

/// <summary>
/// Test script lexical analyzer on OnExcel instr.
/// Very short version.
/// Neither OnSheet, nor FirstRow.
/// One IfThen in ForEachRow.
/// </summary>
[TestClass]
public class ScriptSyntaxAnalyzerOnExcelVeryShortTests
{
    /// <summary>
    /// Compile:, OnExcel, very short version
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
    public void VeryShortOnExcelOk()
    {
        ScriptLineTokensTest lineTok;
        List<ScriptLineTokens> script = new List<ScriptLineTokens>();

        //-build one line of tokens
        lineTok = ScriptLineTokensTest.CreateOnExcel("\"data.xlsx\"");
        script.Add(lineTok);

        // ForEach Row
        lineTok = new ScriptLineTokensTest();
        lineTok.AddTokenName(2, "ForEach", "Row");
        script.Add(lineTok);

        // If A.Cell >10 Then A.Cell=10
        ScriptBuilder.BuidIfACellEq10ThenSetACell(3,script);

        // Next
        lineTok = new ScriptLineTokensTest();
        lineTok.AddTokenName(1, 1, "Next");
        script.Add(lineTok);

        // End OnExcel
        lineTok = new ScriptLineTokensTest();
        lineTok.AddTokenName(1, "End", "OnExcel");
        script.Add(lineTok);


        SyntaxAnalyser sa = new SyntaxAnalyser();

        ExecResult execResult = new ExecResult();
        bool res = sa.Process(execResult, script, out List<InstrBase> listInstr);

        Assert.IsTrue(res);
        Assert.AreEqual(1, listInstr.Count);

        // OnExcel
        Assert.AreEqual(InstrType.OnExcel, listInstr[0].InstrType);
        InstrOnExcel instrOnExcel = listInstr[0] as InstrOnExcel;

        // OnExcel.ListFiles
        Assert.AreEqual(1, instrOnExcel.ListFiles.Count);
        InstrConstValue constExcelFileName = instrOnExcel.ListFiles[0] as InstrConstValue;
        Assert.AreEqual("\"data.xlsx\"", (constExcelFileName.ValueBase as ValueString).Val);

        // check InstrOnSheet
        Assert.AreEqual(1, instrOnExcel.ListSheets.Count);
        InstrOnSheet instrOnSheet = instrOnExcel.ListSheets[0];
        Assert.AreEqual(1, instrOnSheet.SheetNum);

        // check IfThen
        Assert.AreEqual(1, instrOnSheet.ListInstr.Count);
        InstrIfThenElse instrIfThenElse = instrOnSheet.ListInstr[0] as InstrIfThenElse;
        Assert.IsNotNull(instrIfThenElse);

        // check If
        Assert.IsNotNull(instrIfThenElse.InstrIf);
        Assert.IsNotNull(instrIfThenElse.InstrIf.OperandLeft);
        Assert.IsNotNull(instrIfThenElse.InstrIf.OperandRight);
        Assert.IsNotNull(instrIfThenElse.InstrIf.Operator);
        Assert.IsNull(instrIfThenElse.InstrIf.Operand);

        // check If-Operator
        InstrSepComparison instrSepComparison = instrIfThenElse.InstrIf.Operator as InstrSepComparison;
        Assert.IsNotNull(instrSepComparison);
        Assert.AreEqual(SepComparisonOperator.GreaterThan, instrSepComparison.Operator);

        // check If-Operand Left
        TestBuilder.TestInstrColCellFuncValue("If-OperandLeft", instrIfThenElse.InstrIf.OperandLeft, "A", 1);

        // check If-Operand Right
        TestBuilder.TestInstrConstValue("If-OperandRight", instrIfThenElse.InstrIf.OperandRight, 10);

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
        ScriptBuilder.BuidIfACellEq10ThenSetACell(3, script);

        // Next
        lineTok = new ScriptLineTokensTest();
        lineTok.AddTokenName(1, 1, "Next");
        script.Add(lineTok);

        // End OnExcel
        lineTok = new ScriptLineTokensTest();
        lineTok.AddTokenName(1, 1, "End");
        lineTok.AddTokenName(1, 1, "OnExcel");
        script.Add(lineTok);

        SyntaxAnalyser sa = new SyntaxAnalyser();

        ExecResult execResult = new ExecResult();
        bool res = sa.Process(execResult, script, out List<InstrBase> listInstr);

        Assert.IsTrue(res);
        Assert.AreEqual(1, listInstr.Count);

        // OnExcel
        Assert.AreEqual(InstrType.OnExcel, listInstr[0].InstrType);
        InstrOnExcel instrOnExcel = listInstr[0] as InstrOnExcel;

        // OnExcel.ListFiles
        Assert.AreEqual(1, instrOnExcel.ListFiles.Count);
        InstrConstValue constExcelFileName = instrOnExcel.ListFiles[0] as InstrConstValue;
        Assert.AreEqual("\"data.xlsx\"", (constExcelFileName.ValueBase as ValueString).Val);

        // check InstrOnSheet
        Assert.AreEqual(1, instrOnExcel.ListSheets.Count);
        InstrOnSheet instrOnSheet = instrOnExcel.ListSheets[0];
        Assert.AreEqual(1, instrOnSheet.SheetNum);

        // check IfThen
        Assert.AreEqual(1, instrOnSheet.ListInstr.Count);
        InstrIfThenElse instrIfThenElse = instrOnSheet.ListInstr[0] as InstrIfThenElse;
        Assert.IsNotNull(instrIfThenElse);

        // check If
        Assert.IsNotNull(instrIfThenElse.InstrIf);
        Assert.IsNotNull(instrIfThenElse.InstrIf.OperandLeft);
        Assert.IsNotNull(instrIfThenElse.InstrIf.OperandRight);
        Assert.IsNotNull(instrIfThenElse.InstrIf.Operator);
        Assert.IsNull(instrIfThenElse.InstrIf.Operand);

        // check If-Operator
        InstrSepComparison instrSepComparison = instrIfThenElse.InstrIf.Operator as InstrSepComparison;
        Assert.IsNotNull(instrSepComparison);
        Assert.AreEqual(SepComparisonOperator.GreaterThan, instrSepComparison.Operator);

        // check If-Operand Left
        TestBuilder.TestInstrColCellFuncValue("If-OperandLeft", instrIfThenElse.InstrIf.OperandLeft, "A", 1);

        // check If-Operand Right
        TestBuilder.TestInstrConstValue("If-OperandRight", instrIfThenElse.InstrIf.OperandRight, 10);

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
        ScriptBuilder.BuidIfACellEq10ThenSetACell(3, script);

        // If B.Cell >12 Then B.Cell=12
        ScriptBuilder.BuidIfACellEq10ThenSetACell(4, script,"B",">",12,"B",12);

        // Next
        lineTok = new ScriptLineTokensTest();
        lineTok.AddTokenName(1, 1, "Next");
        script.Add(lineTok);

        // End OnExcel
        lineTok = new ScriptLineTokensTest();
        lineTok.AddTokenName(1, 1, "End");
        lineTok.AddTokenName(1, 1, "OnExcel");
        script.Add(lineTok);

        SyntaxAnalyser sa = new SyntaxAnalyser();

        ExecResult execResult = new ExecResult();
        bool res = sa.Process(execResult, script, out List<InstrBase> listInstr);

        Assert.IsTrue(res);
        Assert.AreEqual(1, listInstr.Count);

        // OnExcel
        Assert.AreEqual(InstrType.OnExcel, listInstr[0].InstrType);
        InstrOnExcel instrOnExcel = listInstr[0] as InstrOnExcel;

        // OnExcel.ListFiles
        Assert.AreEqual(1, instrOnExcel.ListFiles.Count);
        InstrConstValue constExcelFileName = instrOnExcel.ListFiles[0] as InstrConstValue;
        Assert.AreEqual("\"data.xlsx\"", (constExcelFileName.ValueBase as ValueString).Val);

        // check InstrOnSheet
        Assert.AreEqual(1, instrOnExcel.ListSheets.Count);
        InstrOnSheet instrOnSheet = instrOnExcel.ListSheets[0];
        Assert.AreEqual(1, instrOnSheet.SheetNum);

        //--check IfThen -> 2 instr!
        Assert.AreEqual(2, instrOnSheet.ListInstr.Count);

        //--check IfThen #1
        InstrIfThenElse instrIfThenElse = instrOnSheet.ListInstr[0] as InstrIfThenElse;
        Assert.IsNotNull(instrIfThenElse);

        // check If
        Assert.IsNotNull(instrIfThenElse.InstrIf);
        Assert.IsNotNull(instrIfThenElse.InstrIf.OperandLeft);
        Assert.IsNotNull(instrIfThenElse.InstrIf.OperandRight);
        Assert.IsNotNull(instrIfThenElse.InstrIf.Operator);
        Assert.IsNull(instrIfThenElse.InstrIf.Operand);

        // check If-Operator
        InstrSepComparison instrSepComparison = instrIfThenElse.InstrIf.Operator as InstrSepComparison;
        Assert.IsNotNull(instrSepComparison);
        Assert.AreEqual(SepComparisonOperator.GreaterThan, instrSepComparison.Operator);

        // check If-Operand Left
        TestBuilder.TestInstrColCellFuncValue("If-OperandLeft", instrIfThenElse.InstrIf.OperandLeft, "A", 1);

        // check If-Operand Right
        TestBuilder.TestInstrConstValue("If-OperandRight", instrIfThenElse.InstrIf.OperandRight, 10);

        // check Then, SetVar -> Left:InstrColCellFunc, Right InstrConstValue: 10
        Assert.IsNotNull(instrIfThenElse.InstrThen);
        Assert.AreEqual(1, instrIfThenElse.InstrThen.ListInstr.Count);

        InstrSetVar instrSetVar = instrIfThenElse.InstrThen.ListInstr[0] as InstrSetVar;
        Assert.IsNotNull(instrSetVar);

        TestBuilder.TestInstrColCellFuncValue("Then-SetVar-OperandLeft", instrSetVar.InstrLeft, "A", 1);
        TestBuilder.TestInstrConstValue("Then-SetVar-OperandRIght", instrSetVar.InstrRight, 10);


        //--check IfThen #2=======
        instrIfThenElse = instrOnSheet.ListInstr[1] as InstrIfThenElse;
        Assert.IsNotNull(instrIfThenElse);

        // check If
        Assert.IsNotNull(instrIfThenElse.InstrIf);
        Assert.IsNotNull(instrIfThenElse.InstrIf.OperandLeft);
        Assert.IsNotNull(instrIfThenElse.InstrIf.OperandRight);
        Assert.IsNotNull(instrIfThenElse.InstrIf.Operator);
        Assert.IsNull(instrIfThenElse.InstrIf.Operand);

        // check If-Operator
        instrSepComparison = instrIfThenElse.InstrIf.Operator as InstrSepComparison;
        Assert.IsNotNull(instrSepComparison);
        Assert.AreEqual(SepComparisonOperator.GreaterThan, instrSepComparison.Operator);

        // check If-Operand Left
        TestBuilder.TestInstrColCellFuncValue("If2-OperandLeft", instrIfThenElse.InstrIf.OperandLeft, "B", 2);

        // check If-Operand Right
        TestBuilder.TestInstrConstValue("If2-OperandRight", instrIfThenElse.InstrIf.OperandRight, 12);

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
        ScriptBuilder.BuidIfACellEq10ThenSetACell(1,lineTok,"A",">",10,"A", 10);
        script.Add(lineTok);
        lineTok.AddTokenName(1, 25, "Next");
        lineTok.AddTokenName(1, 34, "End");
        lineTok.AddTokenName(1, 40, "OnExcel");

        SyntaxAnalyser sa = new SyntaxAnalyser();

        ExecResult execResult = new ExecResult();
        bool res = sa.Process(execResult, script, out List<InstrBase> listInstr);

        Assert.IsTrue(res);
        Assert.AreEqual(1, listInstr.Count);

        // OnExcel
        Assert.AreEqual(InstrType.OnExcel, listInstr[0].InstrType);
        InstrOnExcel instrOnExcel = listInstr[0] as InstrOnExcel;

        // OnExcel.ListFiles
        Assert.AreEqual(1, instrOnExcel.ListFiles.Count);
        InstrConstValue constExcelFileName = instrOnExcel.ListFiles[0] as InstrConstValue;
        Assert.AreEqual("\"data.xlsx\"", (constExcelFileName.ValueBase as ValueString).Val);

        // check InstrOnSheet
        Assert.AreEqual(1, instrOnExcel.ListSheets.Count);
        InstrOnSheet instrOnSheet = instrOnExcel.ListSheets[0];
        Assert.AreEqual(1, instrOnSheet.SheetNum);

        // check IfThen
        Assert.AreEqual(1, instrOnSheet.ListInstr.Count);
        InstrIfThenElse instrIfThenElse = instrOnSheet.ListInstr[0] as InstrIfThenElse;
        Assert.IsNotNull(instrIfThenElse);

        // check If
        Assert.IsNotNull(instrIfThenElse.InstrIf);
        Assert.IsNotNull(instrIfThenElse.InstrIf.OperandLeft);
        Assert.IsNotNull(instrIfThenElse.InstrIf.OperandRight);
        Assert.IsNotNull(instrIfThenElse.InstrIf.Operator);
        Assert.IsNull(instrIfThenElse.InstrIf.Operand);

        // check If-Operator
        InstrSepComparison instrSepComparison = instrIfThenElse.InstrIf.Operator as InstrSepComparison;
        Assert.IsNotNull(instrSepComparison);
        Assert.AreEqual(SepComparisonOperator.GreaterThan, instrSepComparison.Operator);

        // check If-Operand Left
        TestBuilder.TestInstrColCellFuncValue("If-OperandLeft", instrIfThenElse.InstrIf.OperandLeft, "A", 1);

        // check If-Operand Right
        TestBuilder.TestInstrConstValue("If-OperandRight", instrIfThenElse.InstrIf.OperandRight, 10);

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
        lineTok = ScriptLineTokensTest.CreateOnExcel("\"data.xlsx\"");
        script.Add(lineTok);

        //-ForEach Row
        lineTok = new ScriptLineTokensTest();
        lineTok.AddTokenName(2, "ForEach", "Row");
        script.Add(lineTok);

        // If A.Cell >10 Then
        lineTok = new ScriptLineTokensTest();
        lineTok.AddTokenName(3, 1, "If");
        ScriptBuilder.BuidColCellCompValue(3, lineTok, "A", ">", 10 );
        lineTok.AddTokenName(3, 1, "Then");
        script.Add(lineTok);

        // A.Cell=10
        lineTok = new ScriptLineTokensTest();
        ScriptBuilder.BuidColCellCompValue(4, lineTok,"A","=", 10);
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

        SyntaxAnalyser sa = new SyntaxAnalyser();

        ExecResult execResult = new ExecResult();
        bool res = sa.Process(execResult, script, out List<InstrBase> listInstr);

        Assert.IsTrue(res);
        Assert.AreEqual(1, listInstr.Count);

        // OnExcel
        Assert.AreEqual(InstrType.OnExcel, listInstr[0].InstrType);
        InstrOnExcel instrOnExcel = listInstr[0] as InstrOnExcel;

        // OnExcel.ListFiles
        Assert.AreEqual(1, instrOnExcel.ListFiles.Count);
        InstrConstValue constExcelFileName = instrOnExcel.ListFiles[0] as InstrConstValue;
        Assert.AreEqual("\"data.xlsx\"", (constExcelFileName.ValueBase as ValueString).Val);

        // check InstrOnSheet
        Assert.AreEqual(1, instrOnExcel.ListSheets.Count);
        InstrOnSheet instrOnSheet = instrOnExcel.ListSheets[0];
        Assert.AreEqual(1, instrOnSheet.SheetNum);

        // check IfThen
        Assert.AreEqual(1, instrOnSheet.ListInstr.Count);
        InstrIfThenElse instrIfThenElse = instrOnSheet.ListInstr[0] as InstrIfThenElse;
        Assert.IsNotNull(instrIfThenElse);

        // check If
        Assert.IsNotNull(instrIfThenElse.InstrIf);
        Assert.IsNotNull(instrIfThenElse.InstrIf.OperandLeft);
        Assert.IsNotNull(instrIfThenElse.InstrIf.OperandRight);
        Assert.IsNotNull(instrIfThenElse.InstrIf.Operator);
        Assert.IsNull(instrIfThenElse.InstrIf.Operand);

        // check If-Operator
        InstrSepComparison instrSepComparison = instrIfThenElse.InstrIf.Operator as InstrSepComparison;
        Assert.IsNotNull(instrSepComparison);
        Assert.AreEqual(SepComparisonOperator.GreaterThan, instrSepComparison.Operator);

        // check If-Operand Left
        TestBuilder.TestInstrColCellFuncValue("If-OperandLeft", instrIfThenElse.InstrIf.OperandLeft, "A", 1);

        // check If-Operand Right
        TestBuilder.TestInstrConstValue("If-OperandRight", instrIfThenElse.InstrIf.OperandRight, 10);

        // check Then, SetVar -> Left:InstrColCellFunc, Right InstrConstValue: 10
        Assert.IsNotNull(instrIfThenElse.InstrThen);
        Assert.AreEqual(1, instrIfThenElse.InstrThen.ListInstr.Count);

        InstrSetVar instrSetVar = instrIfThenElse.InstrThen.ListInstr[0] as InstrSetVar;
        Assert.IsNotNull(instrSetVar);

        TestBuilder.TestInstrColCellFuncValue("Then-SetVar-OperandLeft", instrSetVar.InstrLeft, "A", 1);
        TestBuilder.TestInstrConstValue("Then-SetVar-OperandRIght", instrSetVar.InstrRight, 10);
    }

}
