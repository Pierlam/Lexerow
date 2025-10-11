using Lexerow.Core.System.Compilator;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Tests._05_Common;
public class ScriptLineTokensTest : ScriptLineTokens
{

    public void AddTokenName(int numLine, string value)
    {
        AddToken(numLine, 1, ScriptTokenType.Name, value);
    }

    public void AddTokenName(int numLine, string value, string value2)
    {
        AddToken(numLine, 1, ScriptTokenType.Name, value);
        AddToken(numLine, value.Length+2, ScriptTokenType.Name, value2);
    }

    /// <summary>
    /// OnExcel "\"data.xlsx\""
    /// </summary>
    /// <param name="numLine"></param>
    /// <param name="excelfile"></param>
    /// <returns></returns>
    public static ScriptLineTokensTest CreateOnExcel(string excelfile)
    {
        var lineTok = new ScriptLineTokensTest();
        lineTok.AddTokenName(1, 1, "OnExcel");
        lineTok.AddTokenString(1, 9, "\"data.xlsx\"");
        return lineTok;
    }

}
