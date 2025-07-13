using Lexerow.Core.System;
using Lexerow.Core.System.Compilator;
using NPOI.SS.Formula.Eval;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Scripts;

/// <summary>
/// Script/source code loader, from a text file.
/// </summary>
public class ScriptLoader
{
    public ScriptLoader()
    {
    }

    /// <summary>
    /// Load a script from a text file.
    /// </summary>
    /// <param name="progName"></param>
    /// <param name="fileName"></param>
    /// <returns></returns>
    public ExecResult LoadScriptFromFile(ExecResult execResult, string fileName, out SourceScript sourceScript)
    {
        sourceScript = null;

        // the file doesn't exists
        if (!File.Exists(fileName))
        {
            execResult.AddError(new ExecResultError(ErrorCode.FileNotFound, fileName));
            return execResult;
        }

        if(!LoadScript(fileName, out sourceScript, out Exception exception))
        {
            execResult.AddError(new ExecResultError(ErrorCode.LoadScriptFileException, exception, fileName));
            return execResult;
        }

        // contains one line at least
        if (sourceScript.Lines.Count == 0) 
        {
            execResult.AddError(new ExecResultError(ErrorCode.LoadScriptFileEmpty, exception, fileName));
            return execResult;
        }

        return execResult;
    }

    private bool LoadScript(string fileName, out SourceScript sourceScript, out Exception exception)
    {
        sourceScript= new SourceScript(fileName);
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
                sourceScript.AddLine(numLine, line);

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
