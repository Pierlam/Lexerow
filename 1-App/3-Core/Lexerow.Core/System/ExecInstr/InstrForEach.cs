using Lexerow.Core.System.Compilator;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;

/// <summary>
/// Instr ForEach. To manage the token found in the script.
/// Used by the parser.
/// </summary>
public class InstrForEach : InstrBase
{
    public InstrForEach(ScriptToken scriptToken) : base(scriptToken)
    {
        InstrType = InstrType.ForEach;
    }
}
