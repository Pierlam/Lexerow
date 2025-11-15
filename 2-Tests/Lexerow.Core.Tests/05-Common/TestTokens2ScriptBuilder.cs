using Lexerow.Core.System.Compilator;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Tests._05_Common;
internal class TestTokens2ScriptBuilder
{
    /// <summary>
    /// Very useful to check that the script is well built.
    /// Generate the initial script from tokens.
    /// </summary>
    /// <param name="scriptTokens"></param>
    /// <returns></returns>
    public static List<string> BuildScript(List<ScriptLineTokens> scriptTokens)
    {
        var script = new List<string>();

        foreach (var lineTokens in scriptTokens)
        {
            string line=string.Empty;

            foreach(var token in lineTokens.ListScriptToken)
            {
                line += token.ToString() + " ";
            }
            script.Add(line);
        }
        return script;
    }
}
