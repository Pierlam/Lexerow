using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System.Compilator;

/// <summary>
/// a script line tokens.
/// One script line split in tokens.
/// </summary>
public class ScriptLineTokens
{
    public int Numline { get; set; } = 0;
    //public ?? srcCode, 

    /// <summary>
    /// The original source code line.
    /// To print when errors occur.
    /// </summary>
    public string ScriptLine { get; set; } = string.Empty;

    /// <summary>
    /// list of tokens of the source code line.
    /// </summary>
    public List<ScriptToken> ListScriptToken { get; set; } = new List<ScriptToken>();

    public void AddTokenName(int numLine, int numCol, string value)
    {
        AddToken(numLine, numCol, ScriptTokenType.Name, value);
    }

    public void AddTokenString(int numLine, int numCol, string value)
    {
        AddToken(numLine, numCol, ScriptTokenType.String, value);
    }

    public void AddTokenInteger(int numLine, int numCol, string value)
    {
        AddToken(numLine, numCol, ScriptTokenType.Integer, value);
    }
    public void AddTokenDouble(int numLine, int numCol, string value)
    {
        AddToken(numLine, numCol, ScriptTokenType.Double, value);
    }

    public void AddTokenSeparator(int numLine, int numCol, string value)
    {
        AddToken(numLine, numCol, ScriptTokenType.Separator, value);
    }

    public void AddToken(int numLine, int numCol, ScriptTokenType type, string value)
    {
        if (string.IsNullOrWhiteSpace(value)) return;

        ScriptToken st = new ScriptToken();
        st.ColNum = numCol;
        st.LineNum = numLine;
        st.ScriptTokenType = type;
        st.Value = value;

        ListScriptToken.Add(st);    
    }

}
