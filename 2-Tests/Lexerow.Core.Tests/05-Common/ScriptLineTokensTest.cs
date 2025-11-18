using Lexerow.Core.System.ScriptDef;

namespace Lexerow.Core.Tests._05_Common;

public class ScriptLineTokensTest : ScriptLineTokens
{
    public void AddTokenInteger(int numLine, int value)
    {
        AddTokenInteger(Numline,1,value);
    }

    public void AddTokenName(int numLine, string value)
    {
        AddToken(numLine, 1, ScriptTokenType.Name, value);
    }

    public void AddTokenName(int numLine, string value, string value2)
    {
        AddToken(numLine, 1, ScriptTokenType.Name, value);
        AddToken(numLine, value.Length + 2, ScriptTokenType.Name, value2);
    }

    /// <summary>
    /// OnExcel "data.xlsx"
    /// </summary>
    /// <param name="numLine"></param>
    /// <param name="excelfile"></param>
    /// <returns></returns>
    public static void CreateOnExcelFileString(int numLine, List<ScriptLineTokens> script,  string excelfileString)
    {
        var line = new ScriptLineTokensTest();
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
    public static ScriptLineTokensTest CreateOnExcelFileName(string excelfileName)
    {
        var lineTok = new ScriptLineTokensTest();
        lineTok.AddTokenName(1, 1, "OnExcel");
        lineTok.AddTokenName(1, 9, excelfileName);
        return lineTok;
    }
}