using Lexerow.Core.System.InstrDef;

namespace Lexerow.Core;

/// <summary>
/// Main data managed by the core.
/// </summary>
public class CoreData
{
    /// <summary>
    /// List of managed programs.
    /// </summary>
    public List<Program> ListProgram { get; set; } = new List<Program>();

    /// <summary>
    /// Get a program script by the name.
    /// </summary>
    /// <param name="name"></param>
    /// <returns></returns>
    public Program GetProgramByName(string name)
    {
        if (string.IsNullOrEmpty(name))
            return null;

        foreach (Program prog in ListProgram)
        {
            if (prog.Script.Name.Equals(name, StringComparison.InvariantCultureIgnoreCase)) return prog;
        }

        return null;
    }
}