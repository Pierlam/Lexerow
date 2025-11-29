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
        InstrType = InstrType.NameObject;
        Name = scriptToken.Value;
    }


    /// <summary>
    /// The name of the object.
    /// Coming from the script token.
    /// </summary>
    public string Name { get; private set; }

    public override string ToString()
    {
        return "NameObject: " + Name;
    }

}