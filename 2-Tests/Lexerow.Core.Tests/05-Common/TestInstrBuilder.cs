using Lexerow.Core.System;
using Lexerow.Core.System.InstrDef;
using Lexerow.Core.System.ScriptDef;
using Lexerow.Core.Utils;

namespace Lexerow.Core.Tests._05_Common;

public class TestInstrBuilder
{
    public static Program CreateProgram()
    {
        return new Program(new Script("scriptName", "fileName"));
    }

    /// <summary>
    /// OnExcel "dataOnExcel1.xlsx"
    ///   ForEach Row
    ///     If..Then
    ///   Next
    /// End OnExcel
    /// </summary>
    /// <param name="fileNameString"></param>
    /// <param name="instrForEach"></param>
    /// <returns></returns>
    public static InstrOnExcel CreateInstrOnExcelFileString(string fileNameString, InstrBase forEachRowInstr)
    {
        var token = CreateScriptTokenString(fileNameString);
        InstrOnExcel instrOnExcel = new InstrOnExcel(token);

        // OnExcel "data.xslx"
        InstrValue instrValue = CreateValueString(fileNameString);
        //instrOnExcel.ListFiles.Add(instrValue);
        instrOnExcel.InstrFiles = instrValue;

        // OnSheet
        var tokenSheet = CreateScriptTokenName("OnSheet");
        instrOnExcel.CreateOnSheet(tokenSheet, 1);

        // ForEach Row instr
        instrOnExcel.CurrOnSheet.ListInstrForEachRow.Add(forEachRowInstr);
        return instrOnExcel;
    }

    /// <summary>
    /// file=SelectFiles("data.xlsx")
    /// OnExcel file
    ///   ForEach Row
    ///     If..Then
    ///   Next
    /// End OnExcel
    /// </summary>
    /// <param name="fileName"></param>
    /// <param name="instrForEach"></param>
    /// <returns></returns>
    public static InstrOnExcel CreateInstrOnExcelFileName(string fileName, InstrBase forEachRowInstr)
    {
        var token = CreateScriptTokenString(fileName);
        InstrOnExcel instrOnExcel = new InstrOnExcel(token);

        // OnExcel file  (varname)
        InstrObjectName instrObjectName = CreateInstrObjectName(fileName);
        instrOnExcel.InstrFiles = instrObjectName;

        // OnSheet
        var tokenSheet = CreateScriptTokenName("OnSheet");
        instrOnExcel.CreateOnSheet(tokenSheet, 1);

        // ForEach Row instr
        instrOnExcel.CurrOnSheet.ListInstrForEachRow.Add(forEachRowInstr);
        return instrOnExcel;
    }

    /// <summary>
    /// If A.Cell >10 Then A.Cell= 10
    /// </summary>
    /// <param name="colNameIf"></param>
    /// <param name="colNumIf"></param>
    /// <param name="sepComp"></param>
    /// <param name="valIf"></param>
    /// <param name="colNameThen"></param>
    /// <param name="colNumThen"></param>
    /// <param name="valThen"></param>
    /// <returns></returns>
    public static InstrIfThenElse CreateInstrIfThen(string colNameIf, int colNumIf, string sepComp, int valIf, string colNameThen, int colNumThen, int valThen)
    {
        // If A.Cell>10
        InstrIf instrIf = CreateInstrIf(colNameIf, colNumIf, sepComp, valIf);

        // Then A.Cell= 10
        InstrThen instrThen = CreateInstrThen(colNameThen, colNumThen, valThen);

        // IfThen
        var tokenIf = CreateScriptTokenName("If");
        InstrIfThenElse instrIfThenElse = new InstrIfThenElse(tokenIf);
        instrIfThenElse.InstrIf = instrIf;
        instrIfThenElse.InstrThen = instrThen;
        return instrIfThenElse;
    }

    /// <summary>
    /// If A.Cell>10
    /// </summary>
    /// <param name="colNameIf"></param>
    /// <param name="colNumIf"></param>
    /// <param name="sepComp"></param>
    /// <returns></returns>
    public static InstrIf CreateInstrIf(string colNameIf, int colNumIf, string sepComp, int val)
    {
        // If
        var tokenIf = CreateScriptTokenName("If");
        InstrIf instrIf = new InstrIf(tokenIf);
        var tokenColNameIf = CreateScriptTokenName("If");
        InstrComparison instrComparison = new InstrComparison(tokenColNameIf);

        // A.Cell
        InstrColCellFunc instrColCellFuncIf = CreateInstrColCellFuncValue(colNameIf, colNumIf);
        instrComparison.OperandLeft = instrColCellFuncIf;

        // sep comp
        var tokenOperatorIf = CreateScriptTokenSep(sepComp);
        instrComparison.Operator = new InstrSepComparison(tokenOperatorIf);

        SepComparisonOperator sep = SepComparisonOperator.GreaterThan;
        if (sepComp == "=") sep = SepComparisonOperator.Equal;
        if (sepComp == "<>") sep = SepComparisonOperator.Different;
        if (sepComp == ">") sep = SepComparisonOperator.GreaterThan;
        if (sepComp == "<") sep = SepComparisonOperator.LessThan;
        if (sepComp == "=>") sep = SepComparisonOperator.GreaterEqualThan;
        if (sepComp == "=<") sep = SepComparisonOperator.LessEqualThan;

        instrComparison.Operator.Operator = sep;

        // val
        var instrRight = CreateInstrValueInt(val);
        instrComparison.OperandRight = instrRight;

        instrIf.InstrBase = instrComparison;
        return instrIf;
    }

    /// <summary>
    /// Then A.Cell= 10
    /// </summary>
    /// <param name="colName"></param>
    /// <param name="colNum"></param>
    /// <param name="val"></param>
    /// <returns></returns>
    public static InstrThen CreateInstrThen(string colName, int colNum, int val)
    {
        var tokenThen = CreateScriptTokenName("Then");
        InstrThen instrThen = new InstrThen(tokenThen);

        // A.Cell
        InstrColCellFunc instrColCellFunc = CreateInstrColCellFuncValue(colName, colNum);

        var tokenSetVar = CreateScriptTokenName("setVar");
        InstrSetVar instrSetVar = new InstrSetVar(tokenSetVar);
        instrSetVar.InstrLeft = instrColCellFunc;
        // 10
        instrSetVar.InstrRight = CreateInstrValueInt(val);

        instrThen.ListInstr.Add(instrSetVar);
        return instrThen;
    }

    /// <summary>
    /// Used for var name.
    /// exp: file
    /// </summary>
    /// <param name="val"></param>
    /// <returns></returns>
    public static InstrObjectName CreateInstrObjectName(string val)
    {
        var script = CreateScriptTokenName(val);
        return new InstrObjectName(script);
    }

    /// <summary>
    /// InstrValue, type: string
    /// exp: "data.xlsx"
    /// </summary>
    /// <param name="val"></param>
    /// <returns></returns>
    public static InstrValue CreateValueString(string val)
    {
        // token
        var token = CreateScriptTokenString(val);
        var str = StringUtils.RemoveStartEndDoubleQuote(val);

        InstrValue instrValue = new InstrValue(token, str);
        instrValue.ValueBase = new ValueString(val);
        return instrValue;
    }

    /// <summary>
    /// InstrValue, type: int
    /// exp: 10
    /// </summary>
    /// <param name="val"></param>
    /// <returns></returns>
    public static InstrValue CreateInstrValueInt(int val)
    {
        // token
        var token = CreateScriptTokenInt(val);

        InstrValue instrValue = new InstrValue(token, val.ToString());
        instrValue.ValueBase = new ValueInt(val);
        return instrValue;
    }

    /// <summary>
    /// Exp: SelectExcel(name)
    /// </summary>
    /// <param name="val"></param>
    /// <returns></returns>
    public static InstrFuncSelectFiles CreateInstrSelectExcelParamObjectName(string val)
    {
        // ObjectName
        var token = CreateScriptTokenName(val);
        var instrObjectName = new InstrObjectName(token);

        // OpenExcel
        InstrFuncSelectFiles instrOpenExcel = new InstrFuncSelectFiles(instrObjectName.FirstScriptToken());
        instrOpenExcel.AddParamSelect(instrObjectName);
        return instrOpenExcel;
    }

    /// <summary>
    /// SelectExcel("data.xslx")
    /// The fileName param is a const value, type string.
    /// </summary>
    /// <param name="fileName"></param>
    /// <returns></returns>
    public static InstrFuncSelectFiles CreateInstrSelectExcelParamString(string fileName)
    {
        InstrValue instrValue = CreateValueString(fileName);
        return CreateInstrSelectExcel(instrValue);
    }

    public static InstrFuncSelectFiles CreateInstrSelectExcel(InstrBase paramFileName)
    {
        InstrFuncSelectFiles instrSelectFiles = new InstrFuncSelectFiles(paramFileName.FirstScriptToken());
        instrSelectFiles.AddParamSelect(paramFileName);
        return instrSelectFiles;
    }

    /// <summary>
    /// SetVar:   a= 12
    ///    InstrLeft:  ObjectName: a
    ///    InstrRight: InstrValue, Int=12
    /// </summary>
    /// <param name="varname"></param>
    /// <param name="value"></param>
    /// <returns></returns>
    public static InstrSetVar CreateInstrSetVarNameValueInt(string varname, int value)
    {
        //-instr left
        InstrObjectName instrObjectName = CreateInstrObjectName(varname);

        //-instr right
        InstrValue instrValue = CreateInstrValueInt(value);

        return CreateInstrSetVar(instrObjectName, instrValue);
    }

    public static InstrSetVar CreateInstrSetVarNameVarName(string varname, string value)
    {
        //-instr left
        InstrObjectName instrObjectName = CreateInstrObjectName(varname);

        //-instr right
        InstrObjectName instrObjectName2 = CreateInstrObjectName(varname);

        return CreateInstrSetVar(instrObjectName, instrObjectName2);
    }

    /// <summary>
    /// Build a SetVar instr.
    /// Format:
    ///   InstrLeft= InstrRight
    /// </summary>
    /// <param name="instrLeft"></param>
    /// <param name="instrRight"></param>
    /// <returns></returns>
    public static InstrSetVar CreateInstrSetVar(InstrBase instrLeft, InstrBase instrRight)
    {
        InstrSetVar instrSetVar = new InstrSetVar(instrLeft.FirstScriptToken());
        instrSetVar.InstrLeft = instrLeft;
        instrSetVar.InstrRight = instrRight;
        return instrSetVar;
    }

    public static string CreateString(string s)
    {
        return "\"" + s + "\"";
    }

    /// <summary>
    /// A.Cell   function=Value
    /// </summary>
    /// <param name="colName"></param>
    /// <param name="colNum"></param>
    /// <returns></returns>
    public static InstrColCellFunc CreateInstrColCellFuncValue(string colName, int colNum)
    {
        var token = CreateScriptTokenName(colName);
        InstrColCellFunc instrColCellFunc = new InstrColCellFunc(token, InstrColCellFuncType.Value, colName, colNum);
        return instrColCellFunc;
    }

    public static ScriptToken CreateScriptTokenName(string val)
    {
        var token = new ScriptToken();
        token.ScriptTokenType = ScriptTokenType.Name;
        token.Value = val;
        return token;
    }

    public static ScriptToken CreateScriptTokenString(string val)
    {
        var token = new ScriptToken();
        token.ScriptTokenType = ScriptTokenType.String;
        // add double quote, start and end
        token.Value = CreateString(val);
        return token;
    }

    /// <summary>
    /// Create token separator:
    ///  =, <>, =>, =<
    /// </summary>
    /// <param name="sep"></param>
    /// <returns></returns>
    public static ScriptToken CreateScriptTokenSep(string sep)
    {
        var token = new ScriptToken();
        token.ScriptTokenType = ScriptTokenType.Separator;
        token.Value = sep;
        return token;
    }

    public static ScriptToken CreateScriptTokenInt(int val)
    {
        var token = new ScriptToken();
        token.ScriptTokenType = ScriptTokenType.Integer;
        token.Value = val.ToString();
        token.ValueInt = val;
        return token;
    }

    public static ScriptToken CreateScriptTokenDouble(double val)
    {
        var token = new ScriptToken();
        token.ScriptTokenType = ScriptTokenType.Double;
        token.Value = val.ToString();
        token.ValueDouble = val;
        return token;
    }
}