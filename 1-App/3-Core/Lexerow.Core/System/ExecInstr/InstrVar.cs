using Lexerow.Core.System.ScriptDef;

namespace Lexerow.Core.System;

/// <summary>
/// Represent a variable name present in the source code
/// exp: a=12
/// sheet.Cell(A,3)=a
///
/// TODO: InstrObjectName exists!!
/// </summary>
public class InstrVar : InstrBase
{
    public InstrVar(ScriptToken scriptToken, string varName) : base(scriptToken)
    {
        InstrType = InstrType.Var;
        VarName = varName;
    }

    public string VarName { get; set; }
}