using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System.Compilator;

/// <summary>
/// obtained by the parser, first stage: lexical analyse.
/// Split the code source line by line to obtains tokens.
/// </summary>
public enum ScriptTokenType
{
    Undefined,

    /// <summary>
    /// Comment or remark
    /// </summary>
    Comment,

    /// <summary>
    /// An item can be: a variable, object, function or method name.
    /// </summary>
    Name,

    /// <summary>
    /// Separator
    /// exp: =,;.< >()+-*/  ...
    /// </summary>
    Separator,

    /// <summary>
    /// Wellformed string,has start and end tags.
    /// </summary>
    String,

    /// <summary>
    /// has the first string tag, but not the last.
    /// </summary>
    StringBadFormed,

    Integer,

    Double,

    /// <summary>
    /// has for example two or more decimal separator.
    /// </summary>
    WrongNumber
}

/// <summary>
/// One script token.
/// Is in a line of script/tokens.
/// Obtained after the lexical analyze (parser).
/// </summary>
public class ScriptToken
{
    /// <summary>
    /// Line Number in the source code.
    /// </summary>
    public int LineNum { get; set; } = 0;

    /// <summary>
    /// Column number in the source code line.
    /// Index first char of the token.
    /// base0
    /// </summary>
    public int ColNum { get; set; }

    /// <summary>
    /// the token value.
    /// </summary>
    public string Value { get; set; }

    public int ValueInt { get; set; } = 0;
    public double ValueDouble { get; set; } = 0;

    public ScriptTokenType ScriptTokenType { get; set; }

    /// <summary>
    /// Can be an integer, double, or a string.
    /// </summary>
    /// <returns></returns>
    public bool IsConstValue()
    {
        // is the token a const value?
        if (ScriptTokenType == ScriptTokenType.Integer || ScriptTokenType == ScriptTokenType.String ||
            ScriptTokenType == ScriptTokenType.StringBadFormed || ScriptTokenType == ScriptTokenType.Double ||
            ScriptTokenType == ScriptTokenType.WrongNumber)
            return true;

        return false;
    }

    public override string ToString()
    {
        return Value;
    }
}
