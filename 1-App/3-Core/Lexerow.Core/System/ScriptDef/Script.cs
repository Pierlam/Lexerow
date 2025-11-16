namespace Lexerow.Core.System.ScriptDef;

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
    /// <summary>
    /// To create an internal one line script easily!
    /// </summary>
    /// <param name="line"></param>
    /// <returns></returns>
    public static Script CreateOneLine(string line)
    {
        if (line == null) return null;
        Script script = new Script("internal", "internal");
        script.AddScriptLine(1, line);
        return script;
    }

    public Script(string scriptName, string fileName)
    {
        Name = scriptName;
        FileName = fileName;
    }

    public string Name { get; set; }

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