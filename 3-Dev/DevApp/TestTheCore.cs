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
        core.Diagnostics.SetLogLevelTrace();
        core.Diagnostics.LogToConsole(true);
        core.Diagnostics.SaveLogTxt(".\\Logs\\logs.txt");
        core.Diagnostics.SaveLogCsv(".\\Logs\\logs.csv");

        string scriptfile = ".\\Scripts\\test2026.lxrw";
        // load the script, compile it and then execute it
        Result result = core.LoadExecScript("script", scriptfile);

        if (result.Res) Console.WriteLine("OK.");
        else Console.WriteLine("ERROR");
    }

    public static void TestCoreCompilCopyHeader()
    {
        LexerowCore core = new LexerowCore();
        //core.ActivityLogEvent += Core_ActivityLogEvent;
        core.Diagnostics.SetLogLevelTrace();
        core.Diagnostics.LogToConsole(true);
        core.Diagnostics.SaveLogTxt(".\\Logs\\logs.txt");

        // create a basic script
        List<string> lines = [
            "file= SelectFiles(\"data.xlsx\")",
            "excelOut= CreateExcel(\"result.xlsx\")",
            "CopyHeader(file, excelOut)",
            ];

        // load the script and compile it
        var result = core.LoadLinesScript("script", lines);
        core.Diagnostics.CloseLogs();

        if (result.Res)
            Console.WriteLine("Result: SUCCESS");
        else
        {
            Console.WriteLine("Result: ERROR, error, line: " + result.ListError[0].LineNum);
        }
    }

}
