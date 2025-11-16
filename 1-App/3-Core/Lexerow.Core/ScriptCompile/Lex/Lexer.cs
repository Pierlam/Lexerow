using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;
using Lexerow.Core.System.ScriptCompile;
using Lexerow.Core.System.ScriptDef;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.ScriptCompile.lex;

/// <summary>
/// script Lexical Analyzer.
/// </summary>
public class Lexer
{
    /// <summary>
    /// Lexical analyze.
    /// Parse a source code script, line by line.
    /// Remove comments, find strings,...
    /// </summary>
    /// <param name="script"></param>
    /// <param name="listScriptLineTokens"></param>
    /// <returns></returns>
    public static bool Process(IActivityLogger logger, ExecResult execResult, Script script, out List<ScriptLineTokens> listScriptLineTokens, LexerConfig lac)
    {
        logger.LogCompilStart(ActivityLogLevel.Important, "Lexer.Process", script.Name);

        ScriptSplitter stringParser = new ScriptSplitter();

        listScriptLineTokens = new List<ScriptLineTokens>();

        int i = -1;
        foreach (ScriptLine scriptLine in script.ScriptLines)
        {
            i++;

            // parse the line, split in tokens
            if(!stringParser.Split(i, scriptLine.Line, lac.Separators, lac.StringSep, lac.CommentTag, out List<ScriptToken> listScriptTokens, out ScriptTokenType lastTokenType))
            {
                if(lastTokenType== ScriptTokenType.WrongNumber)
                {
                    var error= execResult.AddError(ErrorCode.LexerFoundDoubleWrong, scriptLine.NumLine, 0, scriptLine.Line);
                    logger.LogCompilEndError(error, "Lexer.Process", script.Name);
                    return false;
                }
                if (lastTokenType == ScriptTokenType.StringBadFormed)
                {
                    execResult.AddError(new ExecResultError(ErrorCode.LexerFoundSgtringBadFormatted, scriptLine.Line));
                    return false;
                }
                if (lastTokenType == ScriptTokenType.Undefined)
                {
                    execResult.AddError(new ExecResultError(ErrorCode.LexerFoundCharUndefined, scriptLine.Line));
                    return false;
                }
            }

            // save the items: numLine+tokens
            if (listScriptTokens.Count == 0)
                continue;

            ScriptLineTokens scriptLineTokens = new ScriptLineTokens();
            scriptLineTokens.ScriptLine = scriptLine.Line;
            scriptLineTokens.Numline = i;
            scriptLineTokens.ListScriptToken.AddRange(listScriptTokens);
            listScriptLineTokens.Add(scriptLineTokens);

        }

        logger.LogCompilEnd(ActivityLogLevel.Important, "Lexer.Process", script.Name);
        return true;
    }

}
