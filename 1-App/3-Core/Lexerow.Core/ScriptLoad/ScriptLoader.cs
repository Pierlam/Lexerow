using Lexerow.Core.System;
using Lexerow.Core.System.ScriptDef;

namespace Lexerow.Core.ScriptLoad;

/// <summary>
/// Script/source code loader, from a text file.
/// </summary>
public class ScriptLoader
{
    public ScriptLoader()
    {
    }

    /// <summary>
    /// Load a script from lines (list of line).
    /// </summary>
    /// <param name="progName"></param>
    /// <param name="fileName"></param>
    /// <returns></returns>
    public bool LoadScriptFromLines(Result result, string scriptName, List<string> scriptLines, out Script script)
    {
        script = new Script(scriptName, string.Empty);

        int numLine = 1;

        if (scriptLines == null)
        {
            result.AddError(ErrorCode.LoadScriptLinesNull, scriptName);
            return false;
        }
        if (scriptLines.Count == 0)
        {
            result.AddError(ErrorCode.LoadScriptLinesEmpty, scriptName);
            return false;
        }

        foreach (string line in scriptLines)
        {
            script.AddScriptLine(numLine, line);
            numLine++;
        }

        return true;
    }

    /// <summary>
    /// Load a script from a text file.
    /// </summary>
    /// <param name="progName"></param>
    /// <param name="fileName"></param>
    /// <returns></returns>
    public bool LoadScriptFromFile(Result result, string scriptName, string fileName, out Script script)
    {
        script = null;

        // the file doesn't exists
        if (!File.Exists(fileName))
        {
            result.AddError(ErrorCode.FileNotFound, fileName);
            return false;
        }

        if (!LoadScript(scriptName, fileName, out script, out Exception exception))
        {
            result.AddError(ErrorCode.LoadScriptFileException, exception, fileName);
            return false;
        }

        // contains one line at least
        if (script.ScriptLines.Count == 0)
        {
            result.AddError(ErrorCode.LoadScriptFileEmpty, exception, fileName);
            return false;
        }

        return true;
    }

    private bool LoadScript(string scriptName, string fileName, out Script script, out Exception exception)
    {
        script = new Script(scriptName, fileName);
        exception = null;

        int numLine = 1;
        try
        {
            StreamReader sr = new StreamReader(fileName);

            // Read the first line of text
            string line = sr.ReadLine();

            // Continue to read until you reach end of file
            while (line != null)
            {
                script.AddScriptLine(numLine, line);

                //Read the next line
                line = sr.ReadLine();
                numLine++;
            }
            //close the file
            sr.Close();
            return true;
        }
        catch (Exception e)
        {
            exception = e;
            return false;
        }
    }
}