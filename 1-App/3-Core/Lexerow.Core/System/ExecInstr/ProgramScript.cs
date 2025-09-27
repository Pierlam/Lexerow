using Lexerow.Core.System.Compilator;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;

/// <summary>
/// Script compiled, ready to be executed.
/// Program of instructions.
/// </summary>
public class ProgramScript
{
    public ProgramScript(Script script, List<InstrBase> listInstr)
    {
        //Name = name;
        //ScriptFileName= scriptFileName;
        Script =script;
        ListInstr = listInstr;
    }

    public bool IsDefault { get; set; } = false;

    /// <summary>
    /// Is the program created and built directly by adding instructions
    /// or coming from a source code/script and then compiled with success.
    /// </summary>
    //public bool IsBuiltByInstr { get; set; } = true;

    //public string Name { get; set; }

    /// <summary>
    /// Possible to create instructions only in build stage.
    /// </summary>
    public CoreStage Stage { get; set; } = CoreStage.Build;

    /// <summary>
    /// Flat list of tokens to execute. Obtains after Lexical analyse and Syntax analyse.
    /// eol are ExeTokEol.
    /// TODO: RENAME_TO: List<ExeTokBase>
    /// </summary>
    public List<InstrBase> ListInstr { get; set; } = new List<InstrBase>();

    /// <summary>
    /// If the program comes from a source code/script txt file.
    /// </summary>
    //public string ScriptFileName {  get; set; }=string.Empty;

    /// <summary>
    /// Script/source code text lines.
    /// From a text file or from a string stream.
    /// </summary>
    public Script Script { get; private set; }
}
