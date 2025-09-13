using Lexerow.Core.System.Compilator;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;

/// <summary>
/// Base of instructions returning a bool value.
/// Comparison, function
/// but not SetCell or SetVar.
/// </summary>
public class InstrRetBoolBase : ExecTokBase
{
    public InstrRetBoolBase(ScriptToken scriptToken) : base(scriptToken)
    {
    }
}
