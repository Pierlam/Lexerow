using Lexerow.Core;
using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace DevApp;

internal class TestTheCore
{
    static void Core_ActivityLogEvent(object? sender, ActivityLog e)
    {
        //Console.WriteLine(e.When.ToString("yyyy-MM-dd HH:mm:ss - ") + e.Level + " " + e.Module + " " + e.Operation + " " + e.Param);
        if(!string.IsNullOrWhiteSpace(e.Message))
            Console.WriteLine(e.Message);
    }

    /// <summary>
    /// If A.Cell >= 10 Then A.Cell=10
    /// </summary>
    public static void TestCore()
    {
        LexerowCore core = new LexerowCore();
        core.Diagnostics.SetLogLevelTrace();
        //core.Diagnostics.LogToConsole(true);
        core.Diagnostics.SaveLogTxt(".\\Logs\\logs.txt");
        core.Diagnostics.SaveLogCsv(".\\Logs\\logs.csv");

        core.ActivityLogEvent += Core_ActivityLogEvent;


        //string scriptfile = ".\\Scripts\\scriptOk.lxrw";
        //string scriptfile = null;
        //string scriptfile = string.Empty;
        //string scriptfile = ".\\Scripts\\destnotexists.lxrw";
        //string scriptfile = ".\\Scripts\\forEachWrong.lxrw";        
        //string scriptfile = ".\\Scripts\\excelFileNotFound.lxrw";
        string scriptfile = ".\\Scripts\\ifCondMismatch.lxrw";


        // load the script, compile it and then execute it
        Result result = core.LoadExecScript("script", scriptfile);
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
