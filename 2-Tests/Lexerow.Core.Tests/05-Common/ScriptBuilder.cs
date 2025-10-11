using Lexerow.Core.System.Compilator;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Tests._05_Common;

/// <summary>
/// Source code/script builder.
/// </summary>
public class ScriptBuilder
{
    public static Script Build(string l1)
    {
        Script sc= new Script("name", "filename");

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

    /// <summary>
    /// Build this script line:
    ///   If A.Cell >10 Then A.Cell=10
    /// </summary>
    /// <param name="script"></param>
    /// <returns></returns>
    public static ScriptLineTokens BuidIfACellEq10ThenSetACell(int numLine,List<ScriptLineTokens> script)
    {
        var lineTok = new ScriptLineTokens();
        lineTok = BuidIfACellEq10ThenSetACell(numLine,lineTok);
        script.Add(lineTok);
        return lineTok;
    }

    /// <summary>
    /// If A.Cell>10 Then A.Cell=10
    /// </summary>
    /// <param name="lineTok"></param>
    /// <returns></returns>

    public static ScriptLineTokens BuidIfACellEq10ThenSetACell(int numLine, ScriptLineTokens lineTok)
    {
        lineTok.AddTokenName(numLine, 1, "If");
        BuidACellGt10(numLine,lineTok);
        lineTok.AddTokenName(numLine, 1, "Then");
        BuidACellEq10(numLine,lineTok);
        return lineTok;
    }

    /// <summary>
    /// A.Cell>10
    /// </summary>
    /// <param name="lineTok"></param>
    /// <returns></returns>
    public static ScriptLineTokens BuidACellGt10(int numLine, ScriptLineTokens lineTok)
    {
        lineTok.AddTokenName(numLine, 1, "A");
        lineTok.AddTokenSeparator(numLine, 1, ".");
        lineTok.AddTokenName(numLine, 1, "Cell");
        lineTok.AddTokenSeparator(numLine, 1, ">");
        lineTok.AddTokenInteger(numLine, 1, 10);
        return lineTok;
    }

    /// <summary>
    /// A.Cell=10
    /// </summary>
    /// <param name="lineTok"></param>
    /// <returns></returns>
    public static ScriptLineTokens BuidACellEq10(int numLine, ScriptLineTokens lineTok)
    {
        lineTok.AddTokenName(numLine, 1, "A");
        lineTok.AddTokenSeparator(numLine, 1, ".");
        lineTok.AddTokenName(numLine, 1, "Cell");
        lineTok.AddTokenSeparator(numLine, 1, "=");
        lineTok.AddTokenInteger(numLine, 1, 10);
        return lineTok;
    }
}
