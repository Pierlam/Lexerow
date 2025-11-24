using Lexerow.Core.System;
using Lexerow.Core.System.ScriptDef;

namespace Lexerow.Core.Tests._05_Common;

/// <summary>
/// Build tokens from script which is a text.
/// </summary>
public class TestTokensBuilder
{
    public static Script CreateScript(string l1)
    {
        Script sc = new Script("name", "filename");

        sc.AddScriptLine(1, l1);
        return sc;
    }

    public static Script CreateScript(string l1, string l2)
    {
        Script sc = new Script("name", "filename");

        sc.AddScriptLine(1, l1);
        sc.AddScriptLine(2, l2);
        return sc;
    }

    /// <summary>
    /// varname= SelectFiles(fileString)
    /// filestring -> "data.xlsx", should have double quotes!
    /// </summary>
    /// <param name="varName"></param>
    /// <param name="fileString"></param>
    /// <returns></returns>
    public static void AddLineSelectFiles(int numLine, List<ScriptLineTokens> script, string varName, string fileString)
    {
        var line = new ScriptLineTokens();
        line.AddTokenName(numLine, 1, varName);
        line.AddTokenSeparator(numLine, 10, "=");
        line.AddTokenName(numLine, 1, "SelectFiles");
        line.AddTokenSeparator(numLine, 1, "(");
        line.AddTokenString(numLine, 1, fileString);
        line.AddTokenSeparator(numLine, 1, ")");
        script.Add(line);
    }

    /// <summary>
    /// SetVar = intValue
    /// a=10
    /// </summary>
    /// <param name="numLine"></param>
    /// <param name="script"></param>
    /// <param name="varName"></param>
    /// <param name="fileString"></param>
    public static void AddLineSetVarInt(int numLine, List<ScriptLineTokens> script, string varName, int value)
    {
        var line = new ScriptLineTokens();
        line.AddTokenName(numLine, 1, varName);
        line.AddTokenSeparator(numLine, 10, "=");
        line.AddTokenInteger(numLine, 13, value);
        script.Add(line);
    }

    /// <summary>
    /// SetVar = intValue
    /// a=Date(2025,11,23)
    /// a=Date(year,month,day)
    /// </summary>
    /// <param name="numLine"></param>
    /// <param name="script"></param>
    /// <param name="varName"></param>
    /// <param name="fileString"></param>
    public static void AddLineSetVarDate(int numLine, List<ScriptLineTokens> script, string varName, int year, int month, int day)
    {
        var line = new ScriptLineTokens();
        line.AddTokenName(numLine, 1, varName);
        line.AddTokenSeparator(numLine, 2, "=");
        line.AddTokenName(numLine, 4, "Date");
        line.AddTokenSeparator(numLine, 10, "(");
        line.AddTokenInteger(numLine, 12, year);
        line.AddTokenSeparator(numLine, 14, ",");
        line.AddTokenInteger(numLine, 16, month);
        line.AddTokenSeparator(numLine, 20, ",");
        line.AddTokenInteger(numLine, 23, day);
        line.AddTokenSeparator(numLine, 25, ")");
        script.Add(line);
    }

    
    /// <summary>
    /// SetVar = intValue
    /// a=Date(y,11,23)
    /// a=Date(year,month,day)
    /// </summary>
    /// <param name="numLine"></param>
    /// <param name="script"></param>
    /// <param name="varName"></param>
    /// <param name="fileString"></param>
    public static void AddLineSetVarDateVarYear(int numLine, List<ScriptLineTokens> script, string varName, string yearVar, int month, int day)
    {
        var line = new ScriptLineTokens();
        line.AddTokenName(numLine, 1, varName);
        line.AddTokenSeparator(numLine, 2, "=");
        line.AddTokenName(numLine, 4, "Date");
        line.AddTokenSeparator(numLine, 10, "(");
        line.AddTokenName(numLine, 12, yearVar);
        line.AddTokenSeparator(numLine, 14, ",");
        line.AddTokenInteger(numLine, 16, month);
        line.AddTokenSeparator(numLine, 20, ",");
        line.AddTokenInteger(numLine, 23, day);
        line.AddTokenSeparator(numLine, 25, ")");
        script.Add(line);
    }
    /// <summary>
    /// a=-7
    /// </summary>
    /// <param name="numLine"></param>
    /// <param name="script"></param>
    /// <param name="varName"></param>
    /// <param name="value"></param>
    public static void AddLineSetVarMinusInt(int numLine, List<ScriptLineTokens> script, string varName, int value)
    {
        var line = new ScriptLineTokens();
        line.AddTokenName(numLine, 1, varName);
        line.AddTokenSeparator(numLine, 10, "=");
        line.AddTokenSeparator(numLine, 12, "-");
        line.AddTokenInteger(numLine, 13, value);
        script.Add(line);
    }

    /// <summary>
    /// SetVar = intValue
    /// a=10
    /// </summary>
    /// <param name="numLine"></param>
    /// <param name="script"></param>
    /// <param name="varName"></param>
    /// <param name="fileString"></param>
    public static void AddLineSetVarVar(int numLine, List<ScriptLineTokens> script, string varName, string value)
    {
        var line = new ScriptLineTokens();
        line.AddTokenName(numLine, 1, varName);
        line.AddTokenSeparator(numLine, 10, "=");
        line.AddTokenName(numLine, 1, value);
        script.Add(line);
    }

    /// <summary>
    /// A.Cell=10
    /// </summary>
    /// <param name="numLine"></param>
    /// <param name="script"></param>
    /// <param name="colName"></param>
    /// <param name="value"></param>
    public static void AddLineSetVarColCellInt(int numLine, List<ScriptLineTokens> script, string colName, int value)
    {
        var line = new ScriptLineTokens();
        BuidColCellOperInt(numLine++, line, colName, "=", value);
        script.Add(line);
    }

    /// <summary>
    /// OnExcel "data.xlsx"
    /// </summary>
    /// <param name="numLine"></param>
    /// <param name="excelfile"></param>
    /// <returns></returns>
    public static void AddLineOnExcelFileString(int numLine, List<ScriptLineTokens> script, string excelfileString)
    {
        var line = new ScriptLineTokens();
        line.AddTokenName(numLine, 1, "OnExcel");
        line.AddTokenString(numLine, 9, excelfileString);
        script.Add(line);
    }

    /// <summary>
    /// OnExcel file
    /// </summary>
    /// <param name="numLine"></param>
    /// <param name="excelfile"></param>
    /// <returns></returns>
    public static void  CreateOnExcelFileName(int numLine, List<ScriptLineTokens> script, string excelfileName)
    {
        var line = new ScriptLineTokens();
        line.AddTokenName(1, 1, "OnExcel");
        line.AddTokenName(1, 9, excelfileName);
        script.Add(line);
    }

    // FirstRow 3
    public static void AddLineFirstRow(int numLine, List<ScriptLineTokens> script, int firstRowValue)
    {
        var line = new ScriptLineTokens();
        line.AddToken(numLine, 1, ScriptTokenType.Name, "FirstRow");
        line.AddTokenInteger(numLine, 1, firstRowValue);
        script.Add(line);
    }

    // FirstRow varname
    public static void AddLineFirstRowVar(int numLine, List<ScriptLineTokens> script, string varname)
    {
        var line = new ScriptLineTokens();
        line.AddToken(numLine, 1, ScriptTokenType.Name, "FirstRow");
        line.AddTokenName(numLine, 1, varname);
        script.Add(line);
    }

    // ForEach Row
    public static void AddLineForEachRow(int numLine, List<ScriptLineTokens> script)
    {
        var line = new ScriptLineTokens();
        AddTokenName(numLine, line, "ForEach", "Row");
        script.Add(line);
    }

    // Next
    public static void AddLineNext(int numLine, List<ScriptLineTokens> script)
    {
        var line = new ScriptLineTokens();
        line.AddTokenName(numLine, 1, "Next");
        script.Add(line);
    }

    // End OnExcel
    public static void AddLineEndOnExcel(int numLine, List<ScriptLineTokens> script)
    {
        var line = new ScriptLineTokens();
        AddTokenName(numLine, line, "End", "OnExcel");
        script.Add(line);
    }

    /// <summary>
    /// Build this script line:
    ///   If A.Cell >10 Then A.Cell=10
    /// </summary>
    /// <param name="script"></param>
    /// <returns></returns>
    public static ScriptLineTokens BuidIfColCellEqualIntThenSetColCellInt(int numLine, List<ScriptLineTokens> script)
    {
        var line = new ScriptLineTokens();
        line = BuidIfColCellCompIntThenSetColCellInt(numLine, line, "A", ">", 10, "A", 10);
        script.Add(line);
        return line;
    }

    /// <summary>
    /// Build this script line:
    ///   If A.Cell >10 Then A.Cell=10
    /// </summary>
    /// <param name="script"></param>
    /// <returns></returns>
    public static ScriptLineTokens BuidIfColCellCompIntThenSetColCellInt(int numLine, List<ScriptLineTokens> script, string colNameIf, string compIf, int valIf, string colNameThen, int valThen)
    {
        var line = new ScriptLineTokens();
        line = BuidIfColCellCompIntThenSetColCellInt(numLine, line, colNameIf, compIf, valIf, colNameThen, valThen);
        script.Add(line);
        return line;
    }

    /// <summary>
    /// Build this script line:
    ///   If A.Cell >10 Then A.Cell=10
    /// </summary>
    /// <param name="script"></param>
    /// <returns></returns>
    public static ScriptLineTokens BuidIfColCellCompNegIntThenSetColCellInt(int numLine, List<ScriptLineTokens> script, string colNameIf, string compIf, int valIf, string colNameThen, int valThen)
    {
        var line = new ScriptLineTokens();
        line = BuidIfColCellCompNegIntThenSetColCellNegInt(numLine, line, colNameIf, compIf, valIf, colNameThen, valThen);
        script.Add(line);
        return line;
    }

    /// <summary>
    /// Build this script line:
    /// If A.Cell=9 Then A.Cell=Blank
    /// </summary>
    /// <param name="numLine"></param>
    /// <param name="script"></param>
    /// <param name="colNameIf"></param>
    /// <param name="compIf"></param>
    /// <param name="valIf"></param>
    /// <param name="colNameThen"></param>
    /// <param name="keywordThen"></param>
    /// <returns></returns>
    public static ScriptLineTokens BuidIfColCellCompIntThenSetColCellKeyword(int numLine, List<ScriptLineTokens> script, string colNameIf, string compIf, int valIf, string colNameThen, string keywordThen)
    {
        var line = new ScriptLineTokens();
        line.AddTokenName(numLine, 1, "If");
        BuidColCellOperInt(numLine, line, colNameIf, compIf, valIf);
        line.AddTokenName(numLine, 1, "Then");
        BuidColCellEqualKeyword(numLine, line, colNameThen, keywordThen);
        script.Add(line);
        return line;
    }

    /// <summary>
    /// Build this script line:
    ///   If A.Cell=blank Then A.Cell=10
    /// </summary>
    /// <param name="script"></param>
    /// <returns></returns>
    public static ScriptLineTokens BuidIfColCellCompKeywordThenSetColCellInt(int numLine, List<ScriptLineTokens> script, string colNameIf, string compIf, string keyword, string colNameThen, int valThen)
    {
        var line = new ScriptLineTokens();
        line.AddTokenName(numLine, 1, "If");
        BuidColCellOperKeyword(numLine, line, colNameIf, compIf, keyword);
        line.AddTokenName(numLine, 1, "Then");
        BuidColCellEqualInt(numLine, line, colNameThen, valThen);
        script.Add(line);
        return line;
    }

    /// <summary>
    /// If A.Cell>-10 Then A.Cell=-10
    /// </summary>
    /// <param name="line"></param>
    /// <returns></returns>
    public static ScriptLineTokens BuidIfColCellCompNegIntThenSetColCellNegInt(int numLine, ScriptLineTokens line, string colNameIf, string compIf, int valIf, string colNameThen, int valThen)
    {
        line.AddTokenName(numLine, 1, "If");
        BuidColCellOperNegInt(numLine, line, colNameIf, compIf, valIf);
        line.AddTokenName(numLine, 1, "Then");
        BuidColCellOperNegInt(numLine, line, colNameIf, "=", valIf);
        return line;
    }


    /// <summary>
    /// If A.Cell>10 Then A.Cell=10
    /// </summary>
    /// <param name="line"></param>
    /// <returns></returns>
    public static ScriptLineTokens BuidIfColCellCompIntThenSetColCellInt(int numLine, ScriptLineTokens line, string colNameIf, string compIf, int valIf, string colNameThen, int valThen)
    {
        line.AddTokenName(numLine, 1, "If");
        BuidColCellOperInt(numLine, line, colNameIf, compIf, valIf);
        line.AddTokenName(numLine, 1, "Then");
        BuidColCellEqualInt(numLine, line, colNameThen, valThen);
        return line;
    }

    /// <summary>
    /// SetVar or comparison: A.Cell= Blank   or null
    /// </summary>
    /// <param name="numLine"></param>
    /// <param name="line"></param>
    /// <param name="colName"></param>
    /// <param name="valKeyword"></param>
    /// <returns></returns>
    public static ScriptLineTokens BuidColCellEqualKeyword(int numLine, ScriptLineTokens line, string colName, string valKeyword)
    {
        return BuidColCellOperKeyword(numLine, line, colName, "=", valKeyword);
    }

    /// <summary>
    /// SetVar or comparison: A.Cell= 10
    /// </summary>
    /// <param name="line"></param>
    /// <returns></returns>
    public static ScriptLineTokens BuidColCellEqualInt(int numLine, ScriptLineTokens line, string colName, int val)
    {
        return BuidColCellOperInt(numLine, line, colName, "=", val);
    }

    /// <summary>
    /// Comparison or setvar
    /// A.Cell>10, A.Cell=10
    /// </summary>
    /// <param name="line"></param>
    /// <returns></returns>
    public static ScriptLineTokens BuidColCellOperInt(int numLine, ScriptLineTokens line, string colName, string compOrSet, int val)
    {
        line.AddTokenName(numLine, 1, colName);
        line.AddTokenSeparator(numLine, 1, ".");
        line.AddTokenName(numLine, 1, "Cell");
        line.AddTokenSeparator(numLine, 1, compOrSet);
        line.AddTokenInteger(numLine, 1, val);
        return line;
    }

    /// <summary>
    /// Comparison or setvar
    /// A.Cell>-10, A.Cell=-10
    /// </summary>
    /// <param name="line"></param>
    /// <returns></returns>
    public static ScriptLineTokens BuidColCellOperNegInt(int numLine, ScriptLineTokens line, string colName, string compOrSet, int val)
    {
        line.AddTokenName(numLine, 1, colName);
        line.AddTokenSeparator(numLine, 10, ".");
        line.AddTokenName(numLine, 12, "Cell");
        line.AddTokenSeparator(numLine, 13, compOrSet);
        line.AddTokenSeparator(numLine, 14, "-");
        line.AddTokenInteger(numLine, 15, val);
        return line;
    }

    /// <summary>
    /// Comparison or setvar
    /// A.Cell=blank, A.Cell=null
    /// A.Cell<>blank, A.Cell<>null
    /// </summary>
    /// <param name="line"></param>
    /// <returns></returns>
    public static ScriptLineTokens BuidColCellOperKeyword(int numLine, ScriptLineTokens line, string colName, string compOrSet, string valstr)
    {
        line.AddTokenName(numLine, 1, colName);
        line.AddTokenSeparator(numLine, 1, ".");
        line.AddTokenName(numLine, 1, "Cell");
        line.AddTokenSeparator(numLine, 1, compOrSet);
        line.AddTokenName(numLine, 1, valstr);
        return line;
    }

    public static void AddTokenName(int numLine, ScriptLineTokens line, string value, string value2)
    {
        line.AddToken(numLine, 1, ScriptTokenType.Name, value);
        line.AddToken(numLine, value.Length + 2, ScriptTokenType.Name, value2);
    }

}