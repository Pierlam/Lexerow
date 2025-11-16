using Lexerow.Core.System;

namespace Lexerow.Core;

/// <summary>
/// Main data managed by the core.
/// </summary>
public class CoreData
{
    /// <summary>
    /// List of managed programs.
    /// </summary>
    public List<ProgramScript> ListProgram { get; set; } = new List<ProgramScript>();

    /// <summary>
    /// Get a program script by the name.
    /// </summary>
    /// <param name="name"></param>
    /// <returns></returns>
    public ProgramScript GetProgramByName(string name)
    {
        if (string.IsNullOrEmpty(name))
            return null;

        foreach (ProgramScript prog in ListProgram)
        {
            if (prog.Script.Name.Equals(name, StringComparison.InvariantCultureIgnoreCase)) return prog;
        }

        return null;
    }
}