using Lexerow.Core.System.ScriptDef;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;

/// <summary>
/// Object name, can be a var name, function name, ...
/// </summary>
public class InstrObjectName : InstrBase
{
    public InstrObjectName(ScriptToken scriptToken) : base(scriptToken)
    {
        InstrType = InstrType.ObjectName;
        ObjectName = scriptToken.Value;
    }

    // type= undefined, var, function

    public string ObjectName { get; private set; } 
}
