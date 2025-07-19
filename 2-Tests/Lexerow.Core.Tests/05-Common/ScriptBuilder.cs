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
        Script sc= new Script("filename");

        sc.AddScriptLine(1, l1);
        return sc;
    }

    public static Script Build(string l1, string l2)
    {
        Script sc = new Script("filename");

        sc.AddScriptLine(1, l1);
        sc.AddScriptLine(2, l2);
        return sc;
    }

    public static Script Build(string l1, string l2, string l3)
    {
        Script sc = new Script("filename");

        sc.AddScriptLine(1, l1);
        sc.AddScriptLine(2, l2);
        sc.AddScriptLine(3, l3);
        return sc;
    }
}
