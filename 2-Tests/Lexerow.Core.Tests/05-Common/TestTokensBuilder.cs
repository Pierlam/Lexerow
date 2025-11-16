using Lexerow.Core.System.ScriptDef;

namespace Lexerow.Core.Tests._05_Common;

/// <summary>
/// Build tokens from script which is a text.
/// </summary>
public class TestTokensBuilder
{
    public static Script Build(string l1)
    {
        Script sc = new Script("name", "filename");

        sc.AddScriptLine(1, l1);
        return sc;
    }

    public static Script Build(string l1, string l2)
    {
        Script sc = new Script("name", "filename");

        sc.AddScriptLine(1, l1);
        sc.AddScriptLine(2, l2);
        return sc;
    }

    //-line #1
    public static ScriptLineTokens BuildSelectFiles(string varName, string fileString)
    {
        var line = new ScriptLineTokens();
        line.AddTokenName(1, 1, varName);
        line.AddTokenSeparator(1, 1, "=");
        line.AddTokenName(1, 1, "SelectFiles");
        line.AddTokenSeparator(1, 1, "(");
        line.AddTokenString(1, 1, fileString);
        line.AddTokenSeparator(1, 1, ")");
        return line;
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
}