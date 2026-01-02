using Lexerow.Core;
using Lexerow.Core.System;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace DevApp;

internal class TestTheCore
{
    /// <summary>
    /// If A.Cell >= 10 Then A.Cell=10

    /// </summary>
    public static void TestCore()
    {

        LexerowCore core = new LexerowCore();
        core.SetLogLevelTrace();

        core.ActivityLogEvent += Core_ActivityLogEvent1;

        string scriptfile = ".\\Scripts\\test2026.lxrw";
        // load the script, compile it and then execute it
        Result result = core.LoadExecScript("script", scriptfile);

        if (result.Res) Console.WriteLine("OK.");
        else Console.WriteLine("ERROR");
    }

    private static void Core_ActivityLogEvent1(object? sender, Lexerow.Core.System.ActivLog.ActivityLog e)
    {
        string s=string.Empty;

        s= e.When.ToString("dd/MM/yyyy HH:mm:ss fff");

        s += " " + e.Level + " " + e.Module + " " + e.Stage + " " + e.Result+ " " + e.Operation + " " + e.Msg;
        Console.WriteLine(s);
    }
}
