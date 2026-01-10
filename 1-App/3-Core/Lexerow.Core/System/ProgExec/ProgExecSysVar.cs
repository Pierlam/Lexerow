namespace Lexerow.Core.System.ProgExec;

/// <summary>
/// Program execution system/builin variables.
/// exp:
/// $DateFormat=""dd/MM/yyyy"
/// $ForceDateFormat=no
/// </summary>
public class ProgExecSysVar
{
    public ProgExecSysVar(string name, ValueBase valueBase)
    {
        Name = name;
        ValueBase = valueBase;
    }

    public string Name { get; set; }

    public ValueBase ValueBase { get; set; }

    public bool Readonly { get; set; } = true;
}