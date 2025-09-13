using Lexerow.Core.System.Compilator;
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
public class InstrVar : ExecTokBase
{
    public InstrVar(ScriptToken scriptToken, string varName):base(scriptToken)
    {
        ExecTokType = ExecTokType.Var;
        VarName = varName;
    }

    public string VarName { get; set; }
}
