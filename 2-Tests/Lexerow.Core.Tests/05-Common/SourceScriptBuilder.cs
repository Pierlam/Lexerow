using Lexerow.Core.System.Compilator;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Tests._05_Common;
public class SourceScriptBuilder
{
    public static SourceScript Build(string fileName)
    {
        SourceScript sc= new SourceScript(fileName);

        sc.AddLine(1, "# comment");
        sc.AddLine(2, "file= OpenExcel(\"data.xslx\")");

        return sc;
    }

    public static SourceScript BuildOneLineComment(string fileName)
    {
        SourceScript sc = new SourceScript(fileName);
        sc.AddLine(1, "# comment");
        return sc;
    }

}
