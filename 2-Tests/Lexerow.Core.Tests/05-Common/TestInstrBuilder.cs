using Lexerow.Core.System;
using Lexerow.Core.System.ScriptDef;
using Lexerow.Core.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Tests._05_Common;
public class TestInstrBuilder
{
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
        InstrOnExcel instrOnExcel= new InstrOnExcel(token);

        // OnExcel "data.xslx"
        InstrConstValue instrConstValue = BuildInstrConstValueString(fileNameString);
        //instrOnExcel.ListFiles.Add(instrConstValue);
        instrOnExcel.InstrFiles= instrConstValue;

        // OnSheet
        var tokenSheet = CreateScriptTokenName("OnSheet");
        instrOnExcel.CreateOnSheet(tokenSheet,1);

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
        InstrObjectName instrObjectName = BuildInstrObjectName(fileName);
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
        instrIfThenElse.InstrIf= instrIf;
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
        var instrRight = BuildInstrConstValueInt(val);
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
        instrSetVar.InstrLeft= instrColCellFunc;
        // 10
        instrSetVar.InstrRight = BuildInstrConstValueInt(val);

        instrThen.ListInstr.Add(instrSetVar);
        return instrThen;
    }

    /// <summary>
    /// Used for var name.
    /// exp: file
    /// </summary>
    /// <param name="val"></param>
    /// <returns></returns>
    public static InstrObjectName BuildInstrObjectName(string val)
    {
        var script= CreateScriptTokenName(val);
        return new InstrObjectName(script);
    }

    /// <summary>
    /// InstrConstValue, type: string
    /// exp: "data.xlsx"
    /// </summary>
    /// <param name="val"></param>
    /// <returns></returns>
    public static InstrConstValue BuildInstrConstValueString(string val)
    {
        // token
        var token = CreateScriptTokenString(val);
        var str= StringUtils.RemoveStartEndDoubleQuote(val);

        // InstrConstValue
        InstrConstValue instrConstValue= new InstrConstValue(token, str);
        instrConstValue.ValueBase= new ValueString(val);
        return instrConstValue;
    }

    /// <summary>
    /// InstrConstValue, type: int
    /// exp: 10
    /// </summary>
    /// <param name="val"></param>
    /// <returns></returns>
    public static InstrConstValue BuildInstrConstValueInt(int val)
    {
        // token
        var token = CreateScriptTokenInt(val);

        // InstrConstValue
        InstrConstValue instrConstValue = new InstrConstValue(token, val.ToString());
        instrConstValue.ValueBase = new ValueInt(val);
        return instrConstValue;
    }

    /// <summary>
    /// Exp: SelectExcel(name)
    /// </summary>
    /// <param name="val"></param>
    /// <returns></returns>
    public static InstrSelectFiles BuildInstrSelectExcelParamObjectName(string val)
    {
        // ObjectName
        var token = CreateScriptTokenName(val);
        var instrObjectName = new InstrObjectName(token);

        // OpenExcel
        InstrSelectFiles instrOpenExcel = new InstrSelectFiles(instrObjectName.FirstScriptToken());
        instrOpenExcel.AddParamSelect(instrObjectName);
        return instrOpenExcel;
    }


    /// <summary>
    /// SelectExcel("data.xslx")
    /// The fileName param is a const value, type string.
    /// </summary>
    /// <param name="fileName"></param>
    /// <returns></returns>
    public static InstrSelectFiles BuildInstrSelectExcelParamString(string fileName)
    {
        InstrConstValue instrConstValue = BuildInstrConstValueString(fileName);
        return BuildInstrSelectExcel(instrConstValue);
    }

    public static InstrSelectFiles BuildInstrSelectExcel(InstrBase paramFileName)
    {
        InstrSelectFiles instrSelectFiles = new InstrSelectFiles(paramFileName.FirstScriptToken());
        instrSelectFiles.AddParamSelect(paramFileName);
        return instrSelectFiles;
    }

    /// <summary>
    /// Build a SetVar instr.
    /// Format:
    ///   InstrLeft= InstrRight
    /// </summary>
    /// <param name="instrLeft"></param>
    /// <param name="instrRight"></param>
    /// <returns></returns>
    public static InstrSetVar BuildInstrSetVar(InstrBase instrLeft, InstrBase instrRight)
    {
        InstrSetVar instrSetVar = new InstrSetVar(instrLeft.FirstScriptToken());
        instrSetVar.InstrLeft = instrLeft;
        instrSetVar.InstrRight = instrRight;
        return instrSetVar;
    }


    public static string BuildString(string s)
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
        token.Value = BuildString(val);
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
