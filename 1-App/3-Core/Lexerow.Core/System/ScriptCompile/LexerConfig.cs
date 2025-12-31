namespace Lexerow.Core.System.ScriptCompile;

/// <summary>
/// Lexer configuration.
/// used to analyse script lines.
/// </summary>
public class LexerConfig
{
    public string Separators = ".,;=()><+-/*!";

    public char StringSep = '\"';

    public string CommentTag = "#";

    /// <summary>
    /// Used to define system variable like: $DateFormat
    /// </summary>
    public char SystVarStartTag = '$';
}