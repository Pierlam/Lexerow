using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System.Compilator;

/// <summary>
/// One line of a script.
/// </summary>
public class ScriptLine
{
    public ScriptLine(int numLine, string line)
    {
        NumLine = numLine;
        Line = line;
    }

    public int NumLine { get; set; }

    /// <summary>
    /// One line of a script.
    /// </summary>
    public string Line { get; set; }
}

/// <summary>
/// A Script, loaded from a text file. 
/// contains a lsit of lines.
/// </summary>
public class Script
{
    public Script(string fileName)
    {
        FileName = fileName;
    }

    /// <summary>
    /// file name containing the source code/script.
    /// empty is the source code is coming from a string stream.
    /// </summary>
    public string FileName { get; set; }    

    /// <summary>
    /// lines of script/source code
    /// </summary>
    public List<ScriptLine> ScriptLines { get; set; } = new List<ScriptLine>();

    public void AddScriptLine(int numLine, string line)
    {
        if (string.IsNullOrEmpty(line))
            line = string.Empty;

        ScriptLines.Add(new ScriptLine(numLine, line));
    }
}
