using Lexerow.Core.Scripts.Compilator;
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
    public static bool Process(SourceScript srcScript, out List<SourceCodeLineTokens> listSourceCodeLineTokens, LexicalAnalyzerConfig lac)
    {
        StringParser stringParser = new StringParser();

        listSourceCodeLineTokens = new List<SourceCodeLineTokens>();

        int i = -1;
        foreach (SourceScriptLine sourceScriptLine in srcScript.Lines)
        {
            i++;

            // parse the line, split in tokens
            stringParser.Parse(i, sourceScriptLine.Line, lac.Separators, lac.StringSep, lac.CommentTag, out List<SourceCodeToken> items);

            // save the items: numLine+tokens
            if (items.Count == 0)
                continue;

            SourceCodeLineTokens sourceCodeLineTokens = new SourceCodeLineTokens();
            sourceCodeLineTokens.SourceCodeLine = sourceScriptLine.Line;
            sourceCodeLineTokens.Numline = i;
            sourceCodeLineTokens.ListSourceCodeToken = items;
            listSourceCodeLineTokens.Add(sourceCodeLineTokens);
        }
        return true;
    }

}
