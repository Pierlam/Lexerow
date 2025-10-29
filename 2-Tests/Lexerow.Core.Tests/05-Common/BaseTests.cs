using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Tests.Common;
public abstract class BaseTests
{
    public string PathExcelFilesCompil = @"..\..\..\10-Files\";

 //   public string PathExcelFilesRun = @"..\..\..\10-Files\Run\";
    public string PathExcelFilesRun = @".\10-Files\Run\";

    public string PathScriptFiles = @"..\..\..\15-Scripts\";

    /// <summary>
    /// add double quote at the start of the string and also and the end
    /// exp: file.xlsx  -> "file.xlsx"
    /// </summary>
    /// <param name="s"></param>
    /// <returns></returns>
    public string AddDblQuote(string s)
    {
        return "\"" + s + "\"";
    }
}
