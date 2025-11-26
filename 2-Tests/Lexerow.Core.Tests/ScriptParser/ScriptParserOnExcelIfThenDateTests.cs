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
/// Focus on If-Then = Date instr.
/// </summary>
[TestClass]
public class ScriptParserOnExcelIfThenDateTests
{
    /// <summary>
    /// Compile:, OnExcel, very short version
    /// Result: one instruction OnExcel
    /// Implicite: sheet=0, FirstRow=1
    ///
    ///	OnExcel
    ///	  "file.xlsx"
    ///        ForEach Row
    ///            If A.Cell > Date(2020, 10,1) Then A.Cell=Date(2020,2,1)
    ///         Next
    /// End OnExcel
    /// </summary>
    [TestMethod]
    public void OnExcelIfACellGreaterDateOk()
    {
        int numLine = 1;
        List<ScriptLineTokens> scriptTokens = new List<ScriptLineTokens>();
        ScriptLineTokens line;

        //-OnExcel
        line = new ScriptLineTokens();
        line.AddTokenName(numLine++, 1, "OnExcel");
        scriptTokens.Add(line);

        //-"data.xslx"
        line = new ScriptLineTokens();
        line.AddTokenString(numLine++, 1, "\"data.xlsx\"");
        scriptTokens.Add(line);

        //-ForEach Row
        line = new ScriptLineTokens();
        TestTokensBuilder.AddTokenName(numLine++, line, "ForEach", "Row");
        scriptTokens.Add(line);

        // If A.Cell > Date(2020, 10,1) Then A.Cell=Date(2020,2,1)
        TestTokensBuilder.BuidIfColCellCompDateymdThenSetColCellDateymd(numLine++, scriptTokens, "A", ">", 2020, 10, 1, "A", "=", 2020, 2, 1);

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

        // check InstrOnSheet
        Assert.AreEqual(1, instrOnExcel.ListSheets.Count);
        InstrOnSheet instrOnSheet = instrOnExcel.ListSheets[0];
        Assert.AreEqual(1, instrOnSheet.SheetNum);

        //--check IfThen
        Assert.AreEqual(1, instrOnSheet.ListInstrForEachRow.Count);

        //--check IfThen 
        InstrIfThenElse instrIfThenElse = instrOnSheet.ListInstrForEachRow[0] as InstrIfThenElse;
        Assert.IsNotNull(instrIfThenElse);

        // check If
        Assert.IsNotNull(instrIfThenElse.InstrIf);
        InstrComparison instrComparison = instrIfThenElse.InstrIf.InstrBase as InstrComparison;
        Assert.IsNotNull(instrComparison);

        Assert.IsNotNull(instrComparison.OperandLeft);
        Assert.IsNotNull(instrComparison.OperandRight);
        Assert.IsNotNull(instrComparison.Operator);

        // check If-Operator: >
        InstrSepComparison instrSepComparison = instrComparison.Operator as InstrSepComparison;
        Assert.IsNotNull(instrSepComparison);
        Assert.AreEqual(SepComparisonOperator.GreaterThan, instrSepComparison.Operator);

        // check If-Operand Left:A.Cell 
        Assert.IsTrue(TestInstrHelper.TestInstrColCellFuncValue(instrComparison.OperandLeft, "A", 1));

        // check If-Operand Right: Date(2020, 10,1)
        InstrFuncDate instrFuncDate = instrComparison.OperandRight as InstrFuncDate;
        Assert.IsNotNull(instrFuncDate);
        Assert.AreEqual(2020, TestInstrHelper.GetValueInt(instrFuncDate.InstrYear));
        Assert.AreEqual(10, TestInstrHelper.GetValueInt(instrFuncDate.InstrMonth));
        Assert.AreEqual(1, TestInstrHelper.GetValueInt(instrFuncDate.InstrDay));

        // check Then, SetVar -> Left:InstrColCellFunc A.Cell, Right InstrValue: Date(2020,2,1)
        Assert.IsNotNull(instrIfThenElse.InstrThen);
        Assert.AreEqual(1, instrIfThenElse.InstrThen.ListInstr.Count);

        InstrSetVar instrSetVar = instrIfThenElse.InstrThen.ListInstr[0] as InstrSetVar;
        Assert.IsNotNull(instrSetVar);
        Assert.IsTrue(TestInstrHelper.TestInstrColCellFuncValue(instrSetVar.InstrLeft, "A", 1));

        // SetVar InstrRight: Date(2020,2,1)
        instrFuncDate = instrSetVar.InstrRight as InstrFuncDate;
        Assert.IsNotNull(instrFuncDate);
        Assert.AreEqual(2020, TestInstrHelper.GetValueInt(instrFuncDate.InstrYear));
        Assert.AreEqual(2, TestInstrHelper.GetValueInt(instrFuncDate.InstrMonth));
        Assert.AreEqual(1, TestInstrHelper.GetValueInt(instrFuncDate.InstrDay));
    }
}
