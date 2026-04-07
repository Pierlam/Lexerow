using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lxrow;

public class HelpPrinter
{
    public static void PrintHelp(string vers)
    {
        Console.WriteLine();
        Console.WriteLine("=> Lxrow - Execute scripts on Excel file");
        Console.WriteLine();
        Console.WriteLine("Version: " + vers + ", by Pierlam, March 2026");
        Console.WriteLine();
        Console.WriteLine("Lxrow is part of the Lexerow project");
        Console.WriteLine("Website: https://pierlam.github.io/Lexerow/");
        Console.WriteLine();
        Console.WriteLine("Goal:");
        Console.WriteLine("  Lxrow is command line application to execute a script on excel file content to perform actions lile get; set cell values and more.");
        Console.WriteLine("  You want to perform some actions to an excel file, So write a text script.");
        Console.WriteLine("  Then execute your script.");
        Console.WriteLine("  Result is saved into an output excel file.");
        Console.WriteLine();
        Console.WriteLine("Usage:");
        Console.WriteLine("  Lxrow -script <script> -out <result_file>");
        Console.WriteLine();
        Console.WriteLine("Parameters:");
        Console.WriteLine("  -script : Full path of the input script file to execute.");
        Console.WriteLine("  -out   : Full path of the result output file.");
        Console.WriteLine();
        Console.WriteLine("Remark:");
        Console.WriteLine("  If filename contains space, add \" around it.");
        Console.WriteLine();
        Console.WriteLine("Examples:");
        Console.WriteLine("  Lxrow -excel C:\\Input\\script.lxrw -out result.txt");
        Console.WriteLine();
    }
}
