using Lexerow.Core.System.ScriptDef;

namespace Lexerow.Core.System.InstrDef;

/// <summary>
/// Name of an object: variable, function call.
/// Not defined at this point.
/// InstrNameObject
/// </summary>
public class InstrNameObject : InstrBase
{
    public InstrNameObject(ScriptToken scriptToken) : base(scriptToken)
    {
        InstrType = InstrType.ObjectName;
        Name = scriptToken.Value;
    }


    /// <summary>
    /// The name of the object.
    /// Coming from the script token.
    /// </summary>
    public string Name { get; private set; }

    public bool MatchName(string varname)
    {
        if(string.IsNullOrWhiteSpace(varname)) return false;
        if (string.IsNullOrWhiteSpace(Name))return false;
        return Name.Equals(varname, StringComparison.InvariantCultureIgnoreCase);
    }

    public override string ToString()
    {
        return "ObjectName: " + Name;
    }

}