using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System.Compilator;
public class SourceCodeLineTokens
{
    public int Numline { get; set; }
    //public ?? srcCode, 

    /// <summary>
    /// The original source code line.
    /// To print when errors occur.
    /// </summary>
    public string SourceCodeLine { get; set; }

    /// <summary>
    /// list of tokens of the source code line.
    /// </summary>
    public List<SourceCodeToken> ListSourceCodeToken { get; set; }
}
