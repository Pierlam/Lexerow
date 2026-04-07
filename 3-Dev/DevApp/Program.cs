using DevApp;
using DocumentFormat.OpenXml.Spreadsheet;
using Lexerow.Core;
using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;

Console.WriteLine("==>Lexerow Dev:");

void Core_ActivityLogEvent(object? sender, ActivityLog e)
{
    Console.WriteLine(e.When.ToString("yyyy-MM-dd HH:mm:ss - ") + e.Level + " " + e.Module + " " + e.Operation + " " + e.Param);
}

void TestCore()
{
    LexerowCore core = new LexerowCore();
    core.ActivityLogEvent += Core_ActivityLogEvent;
    core.LoadScript("scriptName", "fileName");
}


void TestGetFiles()
{
    //string fileName = @".\Input\*.xlsx";

    string fileName = @".\*.xlsx";
    //string fileName = @".\Input\Test1.xlsx";
    //string fileName = @".\Input\*.bla";
    //string fileName = @".\Bla\*.xlsx";

    string filepath = Path.GetDirectoryName(fileName);
    string fullpath = Path.GetFullPath(fileName);
    string files = Path.GetFileName(fileName);

    if (!Path.Exists(filepath))
    {
        Console.WriteLine($"{fileName} does not exist.");
        return;
    }

    //string filepath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
    DirectoryInfo d = new DirectoryInfo(filepath);
    foreach (var file in d.GetFiles(files))
    {
        Console.WriteLine($"{file.Name}");
        //Directory.Move(file.FullName, filepath + "\\TextFiles\\" + file.Name);
    }
}

//==================================
Main:

//TestCore();

//Test2();

//--check syntax or comparison operator
int a = 12;
if (a >= 10) { }
if (a <= 10) { }

//TestGetFiles();

//DateSandBox.TestReadDates();

//DateSandBox.TestWriteDate();

TestTheCore.TestCore();

//TestTheCore.TestCoreCompilCopyHeader();

Console.WriteLine("\n=> ends.");