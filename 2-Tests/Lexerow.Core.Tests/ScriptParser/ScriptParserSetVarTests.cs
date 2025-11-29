using FakeItEasy;
using Lexerow.Core.ScriptCompile.Parse;
using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;
using Lexerow.Core.System.InstrDef;
using Lexerow.Core.System.InstrDef.Func;
using Lexerow.Core.System.ScriptDef;
using Lexerow.Core.Tests._05_Common;
using Lexerow.Core.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Tests.ScriptParser;

/// <summary>
/// Test script parser basic tests: focus on SetVar instr.
/// </summary>
[TestClass]
public class ScriptParserSetVarTests
{
    /// <summary>
    /// a=12
    /// </summary>
    [TestMethod]
    public void SetaEq12Ok()
    {
        int numLine = 0;
        List<ScriptLineTokens> scriptTokens = new List<ScriptLineTokens>();

        //-a=12
        TestTokensBuilder.AddLineSetVarInt(numLine++, scriptTokens, "a", 12);

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

        // SetVar
        Assert.AreEqual(InstrType.SetVar, prog.ListInstr[0].InstrType);
        InstrSetVar instrSetVar = prog.ListInstr[0] as InstrSetVar;

        // InstrLeft: ObjectName
        InstrNameObject instrObjectName = instrSetVar.InstrLeft as InstrNameObject;
        Assert.IsNotNull(instrObjectName);
        Assert.AreEqual("a", instrObjectName.Name);

        // InstrRight: Value
        InstrValue instrValue = instrSetVar.InstrRight as InstrValue;
        Assert.IsNotNull(instrValue);
        Assert.AreEqual(12, (instrValue.ValueBase as ValueInt).Val);
    }

    /// <summary>
    /// a=-7
    /// </summary>
    [TestMethod]
    public void SetaEqMinus7Ok()
    {
        int numLine = 0;
        List<ScriptLineTokens> scriptTokens = new List<ScriptLineTokens>();

        //-a=-7
        TestTokensBuilder.AddLineSetVarMinusInt(numLine++, scriptTokens, "a", 7);

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

        // SetVar
        Assert.AreEqual(InstrType.SetVar, prog.ListInstr[0].InstrType);
        InstrSetVar instrSetVar = prog.ListInstr[0] as InstrSetVar;

        // InstrLeft: ObjectName
        InstrNameObject instrObjectName = instrSetVar.InstrLeft as InstrNameObject;
        Assert.IsNotNull(instrObjectName);
        Assert.AreEqual("a", instrObjectName.Name);

        // InstrRight: Value
        InstrValue instrValue = instrSetVar.InstrRight as InstrValue;
        Assert.IsNotNull(instrValue);
        Assert.AreEqual(-7, (instrValue.ValueBase as ValueInt).Val);
    }

    /// <summary>
    /// a=12
    /// b=a
    /// </summary>
    [TestMethod]
    public void SetaEq12bEqaOk()
    {
        int numLine = 0;
        List<ScriptLineTokens> scriptTokens = new List<ScriptLineTokens>();

        //-a=12
        TestTokensBuilder.AddLineSetVarInt(numLine++, scriptTokens, "a", 12);

        //-b=a
        TestTokensBuilder.AddLineSetVarVar(numLine++, scriptTokens, "b", "a");

        //==>just to check the content of the script
        var scriptCheck = TestTokens2ScriptBuilder.BuildScript(scriptTokens);

        //==> Parse the script tokens
        Parser parser = new Parser(A.Fake<IActivityLogger>());
        Result result = new Result();
        var prog = TestInstrBuilder.CreateProgram();
        bool res = parser.Process(result, scriptTokens, prog);

        //==> Check the result
        Assert.IsTrue(res);
        Assert.AreEqual(2, prog.ListInstr.Count);

        //--SetVar a=12
        Assert.AreEqual(InstrType.SetVar, prog.ListInstr[0].InstrType);
        InstrSetVar instrSetVar = prog.ListInstr[0] as InstrSetVar;

        // InstrLeft: ObjectName
        InstrNameObject instrObjectName = instrSetVar.InstrLeft as InstrNameObject;
        Assert.IsNotNull(instrObjectName);
        Assert.AreEqual("a", instrObjectName.Name);

        // InstrRight: Value
        InstrValue instrValue = instrSetVar.InstrRight as InstrValue;
        Assert.IsNotNull(instrValue);
        Assert.AreEqual(12, (instrValue.ValueBase as ValueInt).Val);

        //--SetVar b=a
        Assert.AreEqual(InstrType.SetVar, prog.ListInstr[1].InstrType);
        instrSetVar = prog.ListInstr[1] as InstrSetVar;

        // InstrLeft: ObjectName
        instrObjectName = instrSetVar.InstrLeft as InstrNameObject;
        Assert.IsNotNull(instrObjectName);
        Assert.AreEqual("b", instrObjectName.Name);

        // InstrRight: Value
        instrObjectName = instrSetVar.InstrRight as InstrNameObject;
        Assert.IsNotNull(instrObjectName);
        Assert.AreEqual("a", instrObjectName.Name);
    }

    /// <summary>
    /// SetVar a date to the var a.
    /// exp:
    /// a=Date(2025,11,23)
    /// </summary>
    [TestMethod]
    public void SetaEqDateOk()
    {
        int numLine = 0;
        List<ScriptLineTokens> scriptTokens = new List<ScriptLineTokens>();

        TestTokensBuilder.AddLineSetVarDate(numLine++, scriptTokens, "a", 2025,11,23);

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

        //--SetVar a=Date()
        Assert.AreEqual(InstrType.SetVar, prog.ListInstr[0].InstrType);
        InstrSetVar instrSetVar = prog.ListInstr[0] as InstrSetVar;

        // InstrLeft: ObjectName
        InstrNameObject instrObjectName = instrSetVar.InstrLeft as InstrNameObject;
        Assert.IsNotNull(instrObjectName);
        Assert.AreEqual("a", instrObjectName.Name);

        // InstrRight: built-in fct Date
        InstrFuncDate instrFuncDate = instrSetVar.InstrRight as InstrFuncDate;
        Assert.IsNotNull(instrFuncDate);

        Assert.AreEqual(2025, TestInstrHelper.GetValueInt(instrFuncDate.InstrYear));
        Assert.AreEqual(11, TestInstrHelper.GetValueInt(instrFuncDate.InstrMonth));
        Assert.AreEqual(23, TestInstrHelper.GetValueInt(instrFuncDate.InstrDay));
    }

    /// <summary>
    /// SetVar a date to the var a.
    /// exp:
    /// year=2025
    /// a=Date(year,11,23)
    /// </summary>
    [TestMethod]
    public void SetaEqDateVarYearOk()
    {
        int numLine = 0;
        List<ScriptLineTokens> scriptTokens = new List<ScriptLineTokens>();

        // y=2025
        TestTokensBuilder.AddLineSetVarInt(numLine++, scriptTokens, "year", 2025);

        // a=Date(y, 11,23)
        TestTokensBuilder.AddLineSetVarDateVarYear(numLine++, scriptTokens, "a", "year", 11, 23);

        //==>just to check the content of the script
        var scriptCheck = TestTokens2ScriptBuilder.BuildScript(scriptTokens);

        //==> Parse the script tokens
        Parser parser = new Parser(A.Fake<IActivityLogger>());
        Result result = new Result();
        var prog = TestInstrBuilder.CreateProgram();
        bool res = parser.Process(result, scriptTokens, prog);

        //==> Check the result
        Assert.IsTrue(res);
        Assert.AreEqual(2, prog.ListInstr.Count);


        //--SetVar year=2025
        Assert.AreEqual(InstrType.SetVar, prog.ListInstr[0].InstrType);
        InstrSetVar instrSetVar = prog.ListInstr[0] as InstrSetVar;

        //--SetVar a=Date()
        Assert.AreEqual(InstrType.SetVar, prog.ListInstr[1].InstrType);
        instrSetVar = prog.ListInstr[1] as InstrSetVar;

        // InstrLeft: ObjectName
        InstrNameObject instrObjectName = instrSetVar.InstrLeft as InstrNameObject;
        Assert.IsNotNull(instrObjectName);
        Assert.AreEqual("a", instrObjectName.Name);

        // InstrRight: built-in fct Date
        InstrFuncDate instrFuncDate = instrSetVar.InstrRight as InstrFuncDate;
        Assert.IsNotNull(instrFuncDate);

        // the year is a var
        InstrNameObject instrObjectName1 = instrFuncDate.InstrYear as InstrNameObject;
        Assert.IsNotNull(instrObjectName1);
        Assert.AreEqual("year", instrObjectName1.Name);

        Assert.AreEqual(11, TestInstrHelper.GetValueInt(instrFuncDate.InstrMonth));
        Assert.AreEqual(23, TestInstrHelper.GetValueInt(instrFuncDate.InstrDay));
    }

    /// <summary>
    /// SetVar a date to the var a.
    /// exp:
    /// a=Date(2025,11)
    /// </summary>
    [TestMethod]
    public void SetaEqDateDayMissingError()
    {
        int numLine = 0;
        List<ScriptLineTokens> scriptTokens = new List<ScriptLineTokens>();

        TestTokensBuilder.AddLineSetVarDateDayMissing(numLine++, scriptTokens, "a", 2025, 11);

        //==>just to check the content of the script
        var scriptCheck = TestTokens2ScriptBuilder.BuildScript(scriptTokens);

        //==> Parse the script tokens
        Parser parser = new Parser(A.Fake<IActivityLogger>());
        Result result = new Result();
        var prog = TestInstrBuilder.CreateProgram();
        bool res = parser.Process(result, scriptTokens, prog);

        //==> Check the result
        Assert.IsFalse(res);
        Assert.AreEqual(ErrorCode.ParserFctParamCountWrong, result.ListError[0].ErrorCode);
    }

    /// <summary>
    /// SetVar a date to the var a.
    /// exp:
    /// a=Date(2025,25,23)
    /// </summary>
    [TestMethod]
    public void SetaEqDateMonthWrongError()
    {
        int numLine = 0;
        List<ScriptLineTokens> scriptTokens = new List<ScriptLineTokens>();

        TestTokensBuilder.AddLineSetVarDate(numLine++, scriptTokens, "a", 2025, 25, 23);

        //==>just to check the content of the script
        var scriptCheck = TestTokens2ScriptBuilder.BuildScript(scriptTokens);

        //==> Parse the script tokens
        Parser parser = new Parser(A.Fake<IActivityLogger>());
        Result result = new Result();
        var prog = TestInstrBuilder.CreateProgram();
        bool res = parser.Process(result, scriptTokens, prog);

        //==> Check the result
        Assert.IsFalse(res);
        Assert.AreEqual(ErrorCode.ParserFctParamWrong, result.ListError[0].ErrorCode);
    }

    /// <summary>
    /// a=b
    /// </summary>
    [TestMethod]
    public void SetaEqbErrorOk()
    {
        int numLine = 0;
        List<ScriptLineTokens> scriptTokens = new List<ScriptLineTokens>();

        //-a=b
        TestTokensBuilder.AddLineSetVarVar(numLine++, scriptTokens, "a", "b");

        //==>just to check the content of the script
        var scriptCheck = TestTokens2ScriptBuilder.BuildScript(scriptTokens);

        //==> Parse the script tokens
        Parser parser = new Parser(A.Fake<IActivityLogger>());
        Result result = new Result();
        var prog = TestInstrBuilder.CreateProgram();
        bool res = parser.Process(result, scriptTokens, prog);

        //==> Check the result
        Assert.IsFalse(res);
        Assert.AreEqual(ErrorCode.ParserVarWrongRightPart, result.ListError[0].ErrorCode);
    }

    /// <summary>
    /// A.Cell=12
    /// can be used only in OnExcel instr
    /// </summary>
    [TestMethod]
    public void SetACellEq12Error()
    {
        int numLine = 0;
        List<ScriptLineTokens> scriptTokens = new List<ScriptLineTokens>();

        //-A.Cell=12
        TestTokensBuilder.AddLineSetVarColCellInt(numLine++, scriptTokens, "A", 12);

        //==>just to check the content of the script
        var scriptCheck = TestTokens2ScriptBuilder.BuildScript(scriptTokens);

        //==> Parse the script tokens
        Parser parser = new Parser(A.Fake<IActivityLogger>());
        Result result = new Result();
        var prog = TestInstrBuilder.CreateProgram();
        bool res = parser.Process(result, scriptTokens, prog);

        //==> Check the result
        Assert.IsFalse(res);
        Assert.AreEqual(ErrorCode.ParserTokenNotExpected, result.ListError[0].ErrorCode);
    }

}
