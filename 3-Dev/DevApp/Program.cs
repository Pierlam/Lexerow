using Lexerow.Core;
using Lexerow.Core.System.ActivLog;

Console.WriteLine("==>Lexerow Dev:");

void Core_ActivityLogEvent(object? sender, ActivityLog e)
{
    Console.WriteLine("ActivityLogEvent!");
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

///
/// If A.Cell In [ "y", "yes", "ok" ] Then A.Cell= "X"
/// by default, it is case sensitive.
///
/// Case insensitive:
/// If A.Cell In [ "y", "yes", "ok" ] /I Then A.Cell= "X"
///  In /CI -> In case Insensitive
///  In    -> In case Sensitive

//XXXXXXXXXXXXXXXXXXXX-MAIN:

//DevNpoi devNpoi = new DevNpoi();

//string fileName = @"Input\Test1.xlsx";
//devNpoi.TestTypes(fileName);

//string fileName = @"Input\TestTypes.xlsx";
//devNpoi.TestTypes();

//string fileName = @"Input\TestDate.xlsx";
//devNpoi.TestDateTime(fileName);

//devNpoi.TestBlankNull();

//TestCore();

//Test2();

//--check syntax or comparison operator
int a = 12;
if (a >= 10) { }
if (a <= 10) { }

TestGetFiles();

Console.WriteLine("\n=> ends.");