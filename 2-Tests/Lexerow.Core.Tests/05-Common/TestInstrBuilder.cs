using Lexerow.Core.System;
using Lexerow.Core.System.Compilator;
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
    public static InstrOnExcel CreateInstrOnExcelFileString(string fileNameString, InstrBase instrForEach)
    {
        var token = CreateScriptTokenString(fileNameString);
        InstrOnExcel instrOnExcel= new InstrOnExcel(token);

        // OnExcel "data.xslx"
        InstrConstValue instrConstValue = BuildInstrConstValueString(fileNameString);
        instrOnExcel.ListFiles.Add(instrConstValue);

        // OnSheet
        var tokenSheet = CreateScriptTokenName("OnSheet");
        instrOnExcel.CreateOnSheet(tokenSheet,1);

        // ForEach Row instr
        instrOnExcel.CurrOnSheet.ListInstrForEachRow.Add(instrForEach);
        return instrOnExcel;
    }

    /// <summary>
    /// file=OpenExcel("data.xlsx")
    /// OnExcel file
    ///   ForEach Row
    ///     If..Then
    ///   Next
    /// End OnExcel  
    /// </summary>
    /// <param name="fileName"></param>
    /// <param name="instrForEach"></param>
    /// <returns></returns>
    public static InstrOnExcel CreateInstrOnExcelFileName(string fileName, InstrForEach instrForEach)
    {
        var token = CreateScriptTokenString(fileName);
        InstrOnExcel instrOnExcel = new InstrOnExcel(token);

        // OnExcel file
        InstrObjectName instrObjectName = BuildInstrObjectName(fileName);
        //InstrConstValue instrConstValue = BuildInstrConstValueString(fileName);
        instrOnExcel.ListFiles.Add(instrObjectName);

        // OnSheet
        var tokenSheet = CreateScriptTokenName("OnSheet");
        instrOnExcel.CreateOnSheet(tokenSheet, 1);

        // ForEach Row instr
        instrOnExcel.CurrOnSheet.ListInstrForEachRow.Add(instrForEach);
        return instrOnExcel;
    }

    /// <summary>
    /// ForEach Row
    ///    Instr
    /// </summary>
    /// <param name="instrIfThenElse"></param>
    /// <returns></returns>
    public static InstrForEach CreateInstrForEach(InstrIfThenElse instrIfThenElse)
    {
        var token = CreateScriptTokenName("For");
        InstrForEach instrForEach = new InstrForEach(token);
        //instrForEach.ListInstr.Add(instrIfThenElse);
        return instrForEach;
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
    /// Exp: OpenExcel(name)
    /// </summary>
    /// <param name="val"></param>
    /// <returns></returns>
    public static InstrOpenExcel BuildInstrOpenExcelParamObjectName(string val)
    {
        // ObjectName
        var token = CreateScriptTokenName(val);
        var instrObjectName = new InstrObjectName(token);

        // OpenExcel
        InstrOpenExcel instrOpenExcel = new InstrOpenExcel(instrObjectName.FirstScriptToken());
        instrOpenExcel.Param = instrObjectName;
        return instrOpenExcel;
    }


    /// <summary>
    /// OpenExcel("data.xslx")
    /// The fileName param is a const value, type string.
    /// </summary>
    /// <param name="fileName"></param>
    /// <returns></returns>
    public static InstrOpenExcel BuildInstrOpenExcelParamString(string fileName)
    {
        InstrConstValue instrConstValue = BuildInstrConstValueString(fileName);
        return BuildInstrOpenExcel(instrConstValue);
    }

    public static InstrOpenExcel BuildInstrOpenExcel(InstrBase paramFileName)
    {
        InstrOpenExcel instrOpenExcel = new InstrOpenExcel(paramFileName.FirstScriptToken());
        instrOpenExcel.Param = paramFileName;
        return instrOpenExcel;
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
