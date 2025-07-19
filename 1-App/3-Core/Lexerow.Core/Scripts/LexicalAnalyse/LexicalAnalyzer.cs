﻿using Lexerow.Core.Scripts.Compilator;
using Lexerow.Core.System;
using Lexerow.Core.System.Compilator;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Scripts;
public class LexicalAnalyzer
{
    /// <summary>
    /// Lexical analyze.
    /// Parse a source code script, line by line.
    /// Remove comments, find strings,...
    /// </summary>
    /// <param name="script"></param>
    /// <param name="listScriptLineTokens"></param>
    /// <returns></returns>
    public static bool Process(ExecResult execResult, Script script, out List<ScriptLineTokens> listScriptLineTokens, LexicalAnalyzerConfig lac)
    {
        StringParser stringParser = new StringParser();

        listScriptLineTokens = new List<ScriptLineTokens>();

        int i = -1;
        foreach (ScriptLine scriptLine in script.ScriptLines)
        {
            i++;

            // parse the line, split in tokens
            if(!stringParser.Parse(i, scriptLine.Line, lac.Separators, lac.StringSep, lac.CommentTag, out List<ScriptToken> items, out ScriptTokenType lastTokenType))
            {
                if(lastTokenType== ScriptTokenType.DoubleWrong)
                {
                    execResult.AddError(new ExecResultError(ErrorCode.LexAnalyzeFoundDoubleWrong, scriptLine.Line));
                    return false;
                }
                if (lastTokenType == ScriptTokenType.StringBadFormed)
                {
                    execResult.AddError(new ExecResultError(ErrorCode.LexAnalyzeFoundSgtringBadFormatted, scriptLine.Line));
                    return false;
                }
                if (lastTokenType == ScriptTokenType.Undefined)
                {
                    execResult.AddError(new ExecResultError(ErrorCode.LexAnalyzeFoundCharUndefined, scriptLine.Line));
                    return false;
                }
            }

            // save the items: numLine+tokens
            if (items.Count == 0)
                continue;

            ScriptLineTokens scriptLineTokens = new ScriptLineTokens();
            scriptLineTokens.ScriptLine = scriptLine.Line;
            scriptLineTokens.Numline = i;
            scriptLineTokens.ListScriptToken.AddRange(items);
            listScriptLineTokens.Add(scriptLineTokens);

        }
        return true;
    }

}
