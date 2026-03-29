using Lexerow.Core;
using Lexerow.Core.System;
using Lxrow;
using System.Reflection;

Assembly assembly = Assembly.GetExecutingAssembly();
Version version = assembly.GetName().Version;

string vers = version.ToString();

// check arguments
if (args.Length == 0 || args[0].ToLower() == "-help" || args[0].ToLower() == "/?")
{
    HelpPrinter.PrintHelp(vers);
    return;
}


if (args[0].ToLower() == "-debug")
{
    foreach (string s in args)
    {
        Console.WriteLine("arg : " + s);
    }
    // Remove first element by skipping it
    args = args.Skip(1).ToArray();
}


if (!ArgsParser.Parse(args, out ProgParams progParams, out string errMsg))
{
    Console.WriteLine(errMsg);
    return;
}

// add xlsx extension if not exists
//if (Path.GetExtension(progParams.OutputResultFile).Length == 0)
//{
//    progParams.OutputResultFile += ".xlsx";
//}

Console.WriteLine("version " + vers + ", by Pierlam, March 2026");
Console.WriteLine("Ok, Will execute the script: " + progParams.InputScriptFile);

// remove previous result file
//if (File.Exists(progParams.OutputResultFile))
//{
//    Console.WriteLine("Remove previous result file : " + progParams.OutputResultFile);
//    File.Delete(progParams.OutputResultFile);
//}

// check input file exists
if (File.Exists(progParams.InputScriptFile) == false)
{
    Console.WriteLine("Error, script file to execute does not exist : " + progParams.InputScriptFile);
    return;
}

LexerowCore core = new LexerowCore();

try
{
    Result result = core.LoadExecScript("script", progParams.InputScriptFile);
    if (result.Res)
    {
        Console.WriteLine("=> Ok, script executed successfully.");
    }
    else
    {
        Console.WriteLine("=> Error, script execution failed, error count: " + result.ListError.Count);
        Console.WriteLine("=> Warning, script execution failed, warning count: " + result.ListWarning.Count);
        Console.WriteLine("Warning #1: " + result.ListWarning[0].ErrorCode);
        return;
    }
}

catch (Exception ex)
{
    Console.WriteLine("=> Error, exception occurs during script execution : " + ex.Message);
    return;
}


