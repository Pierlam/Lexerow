using Lexerow.Core.System.Compilator;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;

/// <summary>
/// Execution token End of line.
/// </summary>
public class ExecTokEol :ExecTokBase
{
    public ExecTokEol(ScriptToken scriptToken):base(scriptToken)
    {
        ExecTokType = ExecTokType.Eol;
    }


    public override string ToString()
    {
        return "Type: Eol";
    }

}
