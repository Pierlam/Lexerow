using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System.Compilator;
public class LexicalAnalyzerConfig
{
    public string Separators = ".,;=()><+-/*!";
    
    public char StringSep = '\"';

    public string CommentTag = "#";

}
