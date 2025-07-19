using Lexerow.Core.Scripts.Compilator;
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
    /// <param name="srcScript"></param>
    /// <param name="listSourceCodeLineTokens"></param>
    /// <returns></returns>
    public static bool Process(Script srcScript, out List<ScriptLineTokens> listSourceCodeLineTokens, LexicalAnalyzerConfig lac)
    {
        StringParser stringParser = new StringParser();

        listSourceCodeLineTokens = new List<ScriptLineTokens>();

        int i = -1;
        foreach (ScriptLine sourceScriptLine in srcScript.ScriptLines)
        {
            i++;

            // parse the line, split in tokens
            stringParser.Parse(i, sourceScriptLine.Line, lac.Separators, lac.StringSep, lac.CommentTag, out List<ScriptToken> items);

            // save the items: numLine+tokens
            if (items.Count == 0)
                continue;

            ScriptLineTokens sourceCodeLineTokens = new ScriptLineTokens();
            sourceCodeLineTokens.ScriptLine = sourceScriptLine.Line;
            sourceCodeLineTokens.Numline = i;
            sourceCodeLineTokens.ListScriptToken.AddRange(items);
            listSourceCodeLineTokens.Add(sourceCodeLineTokens);

            // pb, string end tag is missing
            //ici();
        }
        return true;
    }

}
