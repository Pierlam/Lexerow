using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System.Compilator;

/// <summary>
/// Lexer configuration.
/// used to analyse script lines.
/// </summary>
public class LexerConfig
{
    public string Separators = ".,;=()><+-/*!";
    
    public char StringSep = '\"';

    public string CommentTag = "#";

}
