using Lexerow.Core.System.Compilator;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;

/// <summary>
/// </summary>
public class ExecTokCloseBracket:ExecTokBase
{
    public ExecTokCloseBracket(ScriptToken scriptToken): base(scriptToken) 
    {
        ExecTokType = ExecTokType.CloseBracket;
    }  
}
