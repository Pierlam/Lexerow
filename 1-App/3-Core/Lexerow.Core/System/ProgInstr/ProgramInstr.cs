using Lexerow.Core.System.Compilator;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;

/// <summary>
/// Program of instructions.
/// </summary>
public class ProgramInstr
{
    public ProgramInstr(string name)
    {
        Name = name;
    }

    public ProgramInstr(string name, string scriptFileName, SourceScript sourceScript, List<InstrBase> listInstr)
    {
        Name = name;
        ScriptFileName= scriptFileName;
        SourceScript =sourceScript;
        ListInstr = listInstr;
    }

    public bool IsDefault { get; set; } = false;

    /// <summary>
    /// Is the program created and built directly by adding instructions
    /// or coming from a source code/script and then compiled with success.
    /// </summary>
    public bool IsBuiltByInstr { get; set; } = true;

    public string Name { get; set; }

    /// <summary>
    /// Possible to create instructions only in build stage.
    /// </summary>
    public CoreStage Stage { get; set; } = CoreStage.Build;

    /// <summary>
    /// list of instructions of the program.
    /// </summary>
    public List<InstrBase> ListInstr { get; set; } = new List<InstrBase>();

    /// <summary>
    /// If the program comes from a source code/script txt file.
    /// </summary>
    public string ScriptFileName {  get; set; }=string.Empty;

    /// <summary>
    /// Script/source code text lines.
    /// Exists only if a script if provided, from a text file or from a string stream.
    /// </summary>
    public SourceScript? SourceScript { get; private set; } = null;
}
