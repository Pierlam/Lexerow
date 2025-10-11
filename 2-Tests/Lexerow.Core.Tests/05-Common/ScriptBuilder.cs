using Lexerow.Core.System.Compilator;
using NPOI.SS.Formula.Functions;
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
        lineTok = BuidIfACellEq10ThenSetACell(numLine, lineTok, "A", ">", 10, "A", 10);
        script.Add(lineTok);
        return lineTok;
    }

    /// <summary>
    /// Build this script line:
    ///   If A.Cell >10 Then A.Cell=10
    /// </summary>
    /// <param name="script"></param>
    /// <returns></returns>
    public static ScriptLineTokens BuidIfACellEq10ThenSetACell(int numLine, List<ScriptLineTokens> script,string colNameIf, string compIf, int valIf, string colNameThen, int valThen)
    {
        var lineTok = new ScriptLineTokens();
        lineTok = BuidIfACellEq10ThenSetACell(numLine, lineTok, colNameIf, compIf, valIf, colNameThen, valThen);
        script.Add(lineTok);
        return lineTok;
    }

    /// <summary>
    /// If A.Cell>10 Then A.Cell=10
    /// </summary>
    /// <param name="lineTok"></param>
    /// <returns></returns>

    public static ScriptLineTokens BuidIfACellEq10ThenSetACell(int numLine, ScriptLineTokens lineTok, string colNameIf, string compIf, int valIf, string colNameThen, int valThen)
    {
        lineTok.AddTokenName(numLine, 1, "If");
        BuidColCellCompValue(numLine,lineTok, colNameIf, compIf, valIf);
        lineTok.AddTokenName(numLine, 1, "Then");
        BuidColCellSetValue(numLine ,lineTok, colNameThen, valThen);
        //BuidACellEq10(numLine,lineTok);
        return lineTok;
    }


    /// <summary>
    /// Comparison
    /// A.Cell>10
    /// </summary>
    /// <param name="lineTok"></param>
    /// <returns></returns>
    public static ScriptLineTokens BuidColCellCompValue(int numLine, ScriptLineTokens lineTok, string colName, string compOrSet, int val)
    {
        lineTok.AddTokenName(numLine, 1, colName);
        lineTok.AddTokenSeparator(numLine, 1, ".");
        lineTok.AddTokenName(numLine, 1, "Cell");
        lineTok.AddTokenSeparator(numLine, 1, compOrSet);
        lineTok.AddTokenInteger(numLine, 1, val);
        return lineTok;
    }

    /// <summary>
    /// SetVar
    /// A.Cell= 10
    /// </summary>
    /// <param name="lineTok"></param>
    /// <returns></returns>
    public static ScriptLineTokens BuidColCellSetValue(int numLine, ScriptLineTokens lineTok, string colName, int val)
    {
        lineTok.AddTokenName(numLine, 1, colName);
        lineTok.AddTokenSeparator(numLine, 1, ".");
        lineTok.AddTokenName(numLine, 1, "Cell");
        lineTok.AddTokenSeparator(numLine, 1, "=");
        lineTok.AddTokenInteger(numLine, 1, val);
        return lineTok;
    }

    /// <summary>
    /// A.Cell=10
    /// </summary>
    /// <param name="lineTok"></param>
    /// <returns></returns>
    //public static ScriptLineTokens BuidACellEq10(int numLine, ScriptLineTokens lineTok)
    //{
    //    lineTok.AddTokenName(numLine, 1, "A");
    //    lineTok.AddTokenSeparator(numLine, 1, ".");
    //    lineTok.AddTokenName(numLine, 1, "Cell");
    //    lineTok.AddTokenSeparator(numLine, 1, "=");
    //    lineTok.AddTokenInteger(numLine, 1, 10);
    //    return lineTok;
    //}


}
