using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System.Compilator;

/// <summary>
/// One line of a source code/script.
/// </summary>
public class SourceScriptLine
{
    public SourceScriptLine(int numLine, string line)
    {
        NumLine = numLine;
        Line = line;
    }

    public int NumLine { get; set; }
    public string Line { get; set; }
}

/// <summary>
/// source code of a script, loaded from a text file. 
/// Split in text lines.
/// </summary>
public class SourceScript
{
    //int _numLine = 0;

    //public List<Error> Errors { get; set; } = new List<Error>();

    public SourceScript(string fileName)
    {
        FileName = fileName;
    }

    /// <summary>
    /// file name containing the source code/script.
    /// empty is the source code is coming from a string stream.
    /// </summary>
    public string FileName { get; set; }    

    public List<SourceScriptLine> Lines { get; set; } = new List<SourceScriptLine>();

    public void AddLine(int numLine, string line)
    {
        if (string.IsNullOrEmpty(line))
            line = string.Empty;

        Lines.Add(new SourceScriptLine(numLine, line));
    }
}
