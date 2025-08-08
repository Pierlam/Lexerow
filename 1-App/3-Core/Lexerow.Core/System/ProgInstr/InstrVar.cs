using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;

/// <summary>
/// Represent a variable name present in the source code
/// exp: a=12
/// sheet.Cell(A,3)=a
/// </summary>
public class InstrVar : InstrBase
{
    public InstrVar(string varName)
    {
        InstrType = InstrType.Var;
        VarName = varName;
    }

    public string VarName { get; set; }
}
