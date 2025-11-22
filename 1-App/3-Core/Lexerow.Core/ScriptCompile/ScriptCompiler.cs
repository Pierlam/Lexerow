using Lexerow.Core.ScriptCompile.lex;
using Lexerow.Core.ScriptCompile.Parse;
using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;
using Lexerow.Core.System.ScriptCompile;
using Lexerow.Core.System.ScriptDef;

namespace Lexerow.Core.ScriptCompile;

/// <summary>
/// Compile script coming from a texte file (or a string) into a list of instructions ready to execute.
/// </summary>
public class ScriptCompiler
{
    private IActivityLogger _logger;

    private CoreData _coreData;
    private LexerConfig lexicalAnalyzerConfig = new LexerConfig();

    /// <summary>
    /// Constructor
    /// </summary>
    /// <param name="coreData"></param>
    public ScriptCompiler(IActivityLogger activityLogger, CoreData coreData)
    {
        _logger = activityLogger;
        _coreData = coreData;
    }

    /// <summary>
    /// Compile the script (source code). this a list of lines.
    /// Generate a list of instructions, ready to be executed.
    /// </summary>
    /// <param name="result"></param>
    /// <param name="script"></param>
    /// <param name="listInstr"></param>
    /// <returns></returns>
    public Result CompileScript(Result result, Script script, out Program programScript)
    {
        _logger.LogCompilStart(ActivityLogLevel.Important, "CompileScript", script.Name);

        // analyse the source code, line by line
        if (!Lexer.Process(_logger, result, script, out List<ScriptLineTokens> listScriptLineTokens, lexicalAnalyzerConfig))
        {
            programScript = null;
            return result;
        }

        // create the program
        programScript = new Program(script);

        Parser parser = new Parser(_logger);
        bool res = parser.Process(result, listScriptLineTokens, programScript);

        if (res)
            _logger.LogCompilEnd(ActivityLogLevel.Important, "CompileScript", script.Name);
        else
            _logger.LogCompilEndError(result.ListError[0], "CompileScript", script.Name);

        return result;
    }
}