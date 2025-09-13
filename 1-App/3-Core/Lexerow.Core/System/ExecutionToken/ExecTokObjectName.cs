using Lexerow.Core.System.Compilator;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;
public class ExecTokObjectName : ExecTokBase
{
    public ExecTokObjectName(ScriptToken scriptToken) : base(scriptToken)
    {
        ExecTokType = ExecTokType.ObjectName;
    }

    // type= undefined, var, function
}
