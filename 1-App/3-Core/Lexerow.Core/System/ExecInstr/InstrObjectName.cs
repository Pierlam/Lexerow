using Lexerow.Core.System.ScriptDef;

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

    public bool MatchName(string varname)
    {
        if(string.IsNullOrWhiteSpace(varname)) return false;
        if (string.IsNullOrWhiteSpace(ObjectName))return false;
        return ObjectName.Equals(varname, StringComparison.InvariantCultureIgnoreCase);
    }
}