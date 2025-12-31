using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;
using Lexerow.Core.System.ScriptCompile;
using Lexerow.Core.System.ScriptDef;

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
    public static bool Process(IActivityLogger logger, Result result, Script script, out List<ScriptLineTokens> listScriptLineTokens, LexerConfig lac)
    {
        logger.LogCompilStart(ActivityLogLevel.Important, "Lexer.Process", script.Name);

        ScriptSplitter stringParser = new ScriptSplitter();

        listScriptLineTokens = new List<ScriptLineTokens>();

        int lineIdx = -1;
        foreach (ScriptLine scriptLine in script.ScriptLines)
        {
            lineIdx++;

            // parse the line, split in tokens
            if (!stringParser.Split(lineIdx, scriptLine.Line, lac.Separators, lac.StringSep, lac.CommentTag, lac.SystVarStartTag, out List<ScriptToken> listScriptTokens, out ScriptTokenType lastTokenType))
            {
                // get the column index in the script line where the error occurs
                int colIdx = 0;
                if (listScriptTokens.Count>0)
                    colIdx = listScriptTokens[listScriptTokens.Count - 1].ColNum;
                
                if (lastTokenType == ScriptTokenType.WrongNumber)
                {
                    var error = result.AddError(ErrorCode.LexerFoundDoubleWrong, lineIdx, colIdx, scriptLine.Line);
                    logger.LogCompilEndError(error, "Lexer.Process", script.Name);
                    return false;
                }
                if (lastTokenType == ScriptTokenType.StringBadFormed)
                {
                    var error = result.AddError(ErrorCode.LexerFoundSgtringBadFormatted, lineIdx, colIdx, scriptLine.Line);
                    logger.LogCompilEndError(error, "Lexer.Process", script.Name);
                    return false;
                }
                if (lastTokenType == ScriptTokenType.WrongSystName)
                {
                    var error = result.AddError(ErrorCode.LexerFoundSystNameWrong, lineIdx, colIdx, scriptLine.Line);
                    logger.LogCompilEndError(error, "Lexer.Process", script.Name);
                    return false;
                }

                if (lastTokenType == ScriptTokenType.Undefined)
                {
                    var error = result.AddError(ErrorCode.LexerFoundCharUndefined, lineIdx, colIdx, scriptLine.Line);
                    logger.LogCompilEndError(error, "Lexer.Process", script.Name);
                    return false;
                }
            }

            // save the items: numLine+tokens
            if (listScriptTokens.Count == 0)
                continue;

            ScriptLineTokens scriptLineTokens = new ScriptLineTokens();
            scriptLineTokens.ScriptLine = scriptLine.Line;
            scriptLineTokens.Numline = lineIdx;
            scriptLineTokens.ListScriptToken.AddRange(listScriptTokens);
            listScriptLineTokens.Add(scriptLineTokens);
        }

        logger.LogCompilEnd(ActivityLogLevel.Important, "Lexer.Process", script.Name);
        return true;
    }
}